//------------------------------------------------------------------------------
// <copyright file="EfModelGeneratorCommand.cs" >
//    
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio;
using System.IO;
using System.Runtime.InteropServices;
using System.Reflection;
using EnvDTE80;
using Microsoft.VisualStudio.Shell.Design;
using EnvDTE;
using System.Collections.Generic;
using System.Text;
using System.Data.Entity.Infrastructure;
using System.Xml;
using System.Data.Entity;
using System.Xml.Linq;
using System.Configuration;

namespace EfModelGenerator
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class EfModelGeneratorCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("219119a1-3233-42d2-b58c-69f775f2b47c");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        private string _fileDirectory;

        private string _fileName;

        private DTE2 _dte2;

        /// <summary>
        /// Initializes a new instance of the <see cref="EfModelGeneratorCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private EfModelGeneratorCommand(Package package, DTE2 dte)
        {

            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;
            _dte2 = dte;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new OleMenuCommand(MenuItemCallback, menuCommandID);
                menuItem.BeforeQueryStatus += menuCommand_BeforeQueryStatus;
                commandService.AddCommand(menuItem);
            }
        }


        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static EfModelGeneratorCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package, DTE2 dte)
        {
            Instance = new EfModelGeneratorCommand(package, dte);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {

            var userConfig = this._dte2.SelectedItems.Item(1).ProjectItem.ContainingProject.GetProjectConfiguration();

            DynamicTypeService typeService;
            IVsSolution solutionService;
            IVsHierarchy projectHierarchy;
            using (ServiceProvider serviceProvider = new ServiceProvider((Microsoft.VisualStudio.OLE.Interop.IServiceProvider)this._dte2.DTE))
            {
                typeService = (DynamicTypeService)serviceProvider.GetService(typeof(DynamicTypeService));
                solutionService = (IVsSolution)serviceProvider.GetService(typeof(SVsSolution));
            }

            int uniqueProjectName = solutionService.GetProjectOfUniqueName(this._dte2.SelectedItems.Item(1).ProjectItem.ContainingProject.UniqueName, out projectHierarchy);

            if (uniqueProjectName != 0)
                throw Marshal.GetExceptionForHR(uniqueProjectName);

            ITypeResolutionService resolutionService = typeService.GetTypeResolutionService(projectHierarchy);
            try
            {
                var itemNamespace = GetTopLevelNamespace(this._dte2.SelectedItems.Item(1).ProjectItem);
                Type type = resolutionService.GetType($"{itemNamespace.FullName}.{_fileName}");

                var dbContextInfo = new DbContextInfo(type, userConfig);

                var filePath = $"{_fileDirectory}\\{_fileName}.edmx";
                using (var writer = new XmlTextWriter(filePath, Encoding.Default))
                {
                    EdmxWriter.WriteEdmx(dbContextInfo.CreateInstance(), writer);
                }

                _dte2.ItemOperations.OpenFile(filePath);
            }
            catch (Exception ex)
            {
                IVsUIShell uiShell = (IVsUIShell)Package.GetGlobalService(typeof(SVsUIShell));
                Guid clsid = Guid.Empty;
                int result;
                uiShell.ShowMessageBox(0, ref clsid, "Ef Model Generator", string.Format(CultureInfo.CurrentCulture, "Inside {0}.Initialize()", this.GetType().FullName),
                                      ex.ToString(), 0, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST, OLEMSGICON.OLEMSGICON_INFO,
                                      0, out result);
            }
        }

        private void menuCommand_BeforeQueryStatus(object sender, EventArgs e)
        {

            var menuCommand = sender as OleMenuCommand;
            if (menuCommand != null)
            {

                menuCommand.Visible = false;
                menuCommand.Enabled = false;

                IVsHierarchy hierarchy = null;
                uint itemid = VSConstants.VSITEMID_NIL;

                if (!IsSingleItemSelected(out hierarchy, out itemid)) return;

                // Get the file path
                var itemFullPath = string.Empty;
                ((IVsProject)hierarchy).GetMkDocument(itemid, out itemFullPath);


                var fileInfo = new FileInfo(itemFullPath);

                _fileName = fileInfo.Name.Replace(fileInfo.Extension, string.Empty);
                _fileDirectory = fileInfo.Directory.FullName;

                var fileEndsDataContext = _fileName.EndsWith("DataContext", StringComparison.InvariantCultureIgnoreCase);

                if (!fileEndsDataContext) return;

                menuCommand.Visible = true;
                menuCommand.Enabled = true;
            }
        }

        private bool IsSingleItemSelected(out IVsHierarchy hierarchy, out uint itemid)
        {
            hierarchy = null;
            itemid = VSConstants.VSITEMID_NIL;
            int hr = VSConstants.S_OK;

            var monitorSelection = Package.GetGlobalService(typeof(SVsShellMonitorSelection)) as IVsMonitorSelection;
            var solution = Package.GetGlobalService(typeof(SVsSolution)) as IVsSolution;
            if (monitorSelection == null || solution == null)
            {
                return false;
            }

            IVsMultiItemSelect multiItemSelect = null;
            IntPtr hierarchyPtr = IntPtr.Zero;
            IntPtr selectionContainerPtr = IntPtr.Zero;

            try
            {
                hr = monitorSelection.GetCurrentSelection(out hierarchyPtr, out itemid, out multiItemSelect, out selectionContainerPtr);

                if (ErrorHandler.Failed(hr) || hierarchyPtr == IntPtr.Zero || itemid == VSConstants.VSITEMID_NIL)
                {
                    // there is no selection
                    return false;
                }

                // multiple items are selected
                if (multiItemSelect != null) return false;

                // there is a hierarchy root node selected, thus it is not a single item inside a project

                if (itemid == VSConstants.VSITEMID_ROOT) return false;

                hierarchy = Marshal.GetObjectForIUnknown(hierarchyPtr) as IVsHierarchy;
                if (hierarchy == null) return false;

                Guid guidProjectID = Guid.Empty;

                if (ErrorHandler.Failed(solution.GetGuidOfProject(hierarchy, out guidProjectID)))
                {
                    return false; // hierarchy is not a project inside the Solution if it does not have a ProjectID Guid
                }

                // if we got this far then there is a single project item selected
                return true;
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

        private CodeElement GetTopLevelNamespace(ProjectItem item)
        {
            var model = item.FileCodeModel;
            foreach (CodeElement element in model.CodeElements)
            {
                if (element.Kind == vsCMElement.vsCMElementNamespace)
                {
                    return element;
                }
            }
            return null;
        }

    }
}
