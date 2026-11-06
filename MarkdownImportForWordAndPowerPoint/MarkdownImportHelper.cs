using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace MarkdownImportForWordAndPowerPoint
{
    /// <summary>
    /// Shared logic of the "Add Markdown file from ..." commands: folder selected in Solution Explorer,
    /// file dialog, conversion, saving the .md file and adding it to the project.
    /// </summary>
    internal static class MarkdownImportHelper
    {
        /// <summary>
        /// Value returned by ShowMessageBox when the user presses "Yes".
        /// </summary>
        private const int IdYes = 6;

        /// <summary>
        /// Folder node selected in Solution Explorer.
        /// </summary>
        private sealed class SelectedFolder
        {
            public string Path { get; set; }

            public IVsHierarchy Hierarchy { get; set; }

            public uint ItemId { get; set; }
        }

        /// <summary>
        /// Runs the whole import flow. Must be called on the UI thread.
        /// </summary>
        /// <param name="package">Owner package.</param>
        /// <param name="messageTitle">Title of the message boxes.</param>
        /// <param name="dialogTitle">Title of the open file dialog.</param>
        /// <param name="filter">Filter of the open file dialog.</param>
        /// <param name="convert">Converts the selected file to Markdown text.</param>
        public static void Run(
            AsyncPackage package,
            string messageTitle,
            string dialogTitle,
            string filter,
            Func<string, string> convert)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // 1. Folder selected in Solution Explorer
            SelectedFolder folder = GetSelectedFolder();
            if (folder == null)
            {
                ShowMessage(package, messageTitle, "Select a single folder in Solution Explorer.");
                return;
            }

            // 2. File dialog to choose the source file
            string sourcePath = PickFile(dialogTitle, filter);
            if (sourcePath == null)
            {
                return; // the user canceled
            }

            // 3. Conversion
            string markdown;
            try
            {
                markdown = convert(sourcePath);
            }
            catch (Exception ex)
            {
                ShowMessage(package, messageTitle, "The conversion failed: " + ex.Message);
                return;
            }

            if (string.IsNullOrWhiteSpace(markdown))
            {
                ShowMessage(package, messageTitle, "No convertible text was found in the selected file.");
                return;
            }

            // 4. Target .md path (same name as the source file, .md extension)
            string markdownPath = Path.Combine(folder.Path, Path.GetFileNameWithoutExtension(sourcePath) + ".md");
            bool alreadyExisted = File.Exists(markdownPath);

            if (alreadyExisted &&
                !AskYesNo(package, messageTitle, "The file '" + Path.GetFileName(markdownPath) + "' already exists in the folder. Do you want to overwrite it?"))
            {
                return;
            }

            // 5. Save (UTF-8 without BOM)
            try
            {
                File.WriteAllText(markdownPath, markdown, new UTF8Encoding(false));
            }
            catch (Exception ex)
            {
                ShowMessage(package, messageTitle, "The file could not be saved: " + ex.Message);
                return;
            }

            // 6. Add the file to the project, under the selected folder
            bool added = AddFileToProject(folder, markdownPath);

            if (added || alreadyExisted)
            {
                ShowMessage(package, messageTitle, (alreadyExisted ? "Markdown file updated: " : "Markdown file created: ") + markdownPath);
            }
            else
            {
                ShowMessage(
                    package,
                    messageTitle,
                    "The Markdown file was created, but it could not be added to the project automatically:" +
                    Environment.NewLine + markdownPath);
            }
        }

        /// <summary>
        /// Returns the folder node selected in Solution Explorer (path, hierarchy and item ID),
        /// or null if the selection is not a single folder.
        /// </summary>
        private static SelectedFolder GetSelectedFolder()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var monitorSelection = Package.GetGlobalService(typeof(SVsShellMonitorSelection)) as IVsMonitorSelection;
            if (monitorSelection == null)
            {
                return null;
            }

            IntPtr hierarchyPtr = IntPtr.Zero;
            IntPtr selectionContainerPtr = IntPtr.Zero;

            try
            {
                int hr = monitorSelection.GetCurrentSelection(
                    out hierarchyPtr,
                    out uint itemId,
                    out IVsMultiItemSelect multiItemSelect,
                    out selectionContainerPtr);

                // Invalid or multiple selection
                if (ErrorHandler.Failed(hr) || hierarchyPtr == IntPtr.Zero || multiItemSelect != null)
                {
                    return null;
                }

                // Exclude the project/solution node and special selections: a folder node is required
                if (itemId == VSConstants.VSITEMID_NIL ||
                    itemId == VSConstants.VSITEMID_ROOT ||
                    itemId == VSConstants.VSITEMID_SELECTION)
                {
                    return null;
                }

                var hierarchy = Marshal.GetObjectForIUnknown(hierarchyPtr) as IVsHierarchy;
                if (hierarchy == null)
                {
                    return null;
                }

                if (ErrorHandler.Failed(hierarchy.GetCanonicalName(itemId, out string path)) ||
                    string.IsNullOrEmpty(path))
                {
                    return null;
                }

                path = path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

                // It must be an existing folder on disk
                if (!Directory.Exists(path))
                {
                    return null;
                }

                return new SelectedFolder { Path = path, Hierarchy = hierarchy, ItemId = itemId };
            }
            finally
            {
                if (selectionContainerPtr != IntPtr.Zero)
                {
                    Marshal.Release(selectionContainerPtr);
                }

                if (hierarchyPtr != IntPtr.Zero)
                {
                    Marshal.Release(hierarchyPtr);
                }
            }
        }

        /// <summary>
        /// Adds an existing file to the project that owns the selected folder, under that folder.
        /// Returns true if the file is in the project, false if the project does not accept it.
        /// </summary>
        private static bool AddFileToProject(SelectedFolder folder, string filePath)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var project = folder.Hierarchy as IVsProject;
            if (project == null)
            {
                return false;
            }

            try
            {
                // Nothing to do if the file is already part of the project (for example when overwriting)
                var priority = new VSDOCUMENTPRIORITY[1];
                if (ErrorHandler.Succeeded(project.IsDocumentInProject(filePath, out int found, priority, out uint existingItemId)) &&
                    found != 0)
                {
                    return true;
                }

                var result = new VSADDRESULT[1];
                int hr = project.AddItem(
                    folder.ItemId,
                    VSADDITEMOPERATION.VSADDITEMOP_OPENFILE,
                    null,
                    1,
                    new[] { filePath },
                    IntPtr.Zero,
                    result);

                return ErrorHandler.Succeeded(hr) && result[0] == VSADDRESULT.ADDRESULT_Success;
            }
            catch (COMException)
            {
                return false;
            }
        }

        /// <summary>
        /// Shows the open file dialog and returns the selected file, or null if canceled.
        /// </summary>
        private static string PickFile(string title, string filter)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = title,
                Filter = filter,
                CheckFileExists = true,
                Multiselect = false
            };

            bool? result = dialog.ShowDialog();
            return result == true ? dialog.FileName : null;
        }

        private static void ShowMessage(AsyncPackage package, string title, string message)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            VsShellUtilities.ShowMessageBox(
                package,
                message,
                title,
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        private static bool AskYesNo(AsyncPackage package, string title, string message)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            int result = VsShellUtilities.ShowMessageBox(
                package,
                message,
                title,
                OLEMSGICON.OLEMSGICON_QUERY,
                OLEMSGBUTTON.OLEMSGBUTTON_YESNO,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_SECOND);

            return result == IdYes;
        }
    }
}
