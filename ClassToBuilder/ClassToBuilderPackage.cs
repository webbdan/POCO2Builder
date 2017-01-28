using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using Microsoft.Win32;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using System.IO;
using EnvDTE;
using System.Collections.Generic;
using System.Xml;
using System.CodeDom.Compiler;
using Microsoft.CSharp;
using System.Reflection;
using System.Xml.Linq;
using System.Linq;
using System.Windows.Forms;

namespace BJSS.ClassToBuilder
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidClassToBuilderPkgString)]
    public sealed class ClassToBuilderPackage : Package
    {
        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public ClassToBuilderPackage()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));
        }

        private static readonly string IsBuilderFile = "IsBuilderFile";

        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation
        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            Debug.WriteLine (string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
            base.Initialize();

            // Add our command handlers for menu (commands must exist in the .vsct file)
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if ( null != mcs )
            {
                // Create the command for the menu item.
                CommandID menuCommandID = new CommandID(GuidList.guidClassToBuilderCmdSet, (int)PkgCmdIDList.cmdidCBuilder);
                //MenuCommand menuItem = new MenuCommand(MenuItemCallback, menuCommandID );
                var menuCommand = new OleMenuCommand(OnAddBuilderCommand, menuCommandID);
                menuCommand.BeforeQueryStatus += OnBeforeQueryStatusAddBuilderCommand;

                mcs.AddCommand(menuCommand);
            }
        }
        #endregion

         private void OnBeforeQueryStatusAddBuilderCommand(object sender, EventArgs e)
        {
            // get the menu that fired the event
            var menuCommand = sender as OleMenuCommand;
            if (menuCommand != null)
            {
                // start by assuming that the menu will not be shown
                menuCommand.Visible = false;
                menuCommand.Enabled = false;

                IVsHierarchy hierarchy = null;
                uint itemid = VSConstants.VSITEMID_NIL;

                if (!IsSingleProjectItemSelection(out hierarchy, out itemid)) return;

                var vsProject = (IVsProject)hierarchy;
                if (!ProjectSupportsBuilders(vsProject)) return;

                if (!ItemSupportsBuilders(vsProject, itemid)) return;

                menuCommand.Visible = true;
                menuCommand.Enabled = true;
            }
        }

        private void OnAddBuilderCommand(object sender, EventArgs e)
        {
            IVsHierarchy hierarchy = null;
            uint itemid = VSConstants.VSITEMID_NIL;

            if (!IsSingleProjectItemSelection(out hierarchy, out itemid)) return;

            var vsProject = (IVsProject)hierarchy;
            if (!ProjectSupportsBuilders(vsProject)) return;

            string projectFullPath = null;
            if (ErrorHandler.Failed(vsProject.GetMkDocument(VSConstants.VSITEMID_ROOT, out projectFullPath))) return;

            var buildPropertyStorage = vsProject as IVsBuildPropertyStorage;
            if (buildPropertyStorage == null) return;
            
            // get the name of the item
            string itemFullPath = null;
            if (ErrorHandler.Failed(vsProject.GetMkDocument(itemid, out itemFullPath))) return;

            // Save the project file
            var solution = (IVsSolution)Package.GetGlobalService(typeof(SVsSolution));
            int hr = solution.SaveSolutionElement((uint)__VSSLNSAVEOPTIONS.SLNSAVEOPT_SaveIfDirty, hierarchy, 0);
            if (hr < 0)
            {
                throw new COMException(string.Format("Failed to add project item {0} {1}", itemFullPath, GetErrorInfo()), hr);
            }

            var selectedProjectItem = GetProjectItemFromHierarchy(hierarchy, itemid);

            CodeElements elts = null;
            elts = selectedProjectItem.FileCodeModel.CodeElements;
            CodeElement elt = null;
            int i = 0;
            myBuilder b = new myBuilder();
            for (i = 1; i <= selectedProjectItem.FileCodeModel.CodeElements.Count; i++)
            {
                elt = elts.Item(i);
                b = CollapseElt(elt, elts, i, new myBuilder());
                if (b.buildername != null)
                {
                    break;
                }
            }
            string content = @"using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
";
            content += b.body;

            Project project = null;
            if (selectedProjectItem != null)
            {
                string itemFolder = Path.GetDirectoryName(itemFullPath);
                string itemFilename = Path.GetFileNameWithoutExtension(itemFullPath);
                string itemExtension = Path.GetExtension(itemFullPath);

                string itemName = itemFilename + "Builder.cs";

                var newItemFullPath = itemFolder; 

                AddBuilderFile(content, itemName, itemFolder, hierarchy);
                AddBuilderToSolution(selectedProjectItem.ContainingProject, selectedProjectItem, itemName, itemFolder, hierarchy);

                // also add the web.oncheckin.config file as well if it's there.
                //AddBuilderToSolution(selectedProjectItem, "web.oncheckin.config", itemFolder, hierarchy);

                // now force the storage type.
                //uint addedFileId;
                //hierarchy.ParseCanonicalName(newItemFullPath, out addedFileId);
                //buildPropertyStorage.SetItemAttribute(addedFileId, IsBuilderFile, "True");
            }
        }

        public myBuilder CollapseElt( CodeElement elt, CodeElements elts, long loc, myBuilder builder ) 
        {
            EditPoint epStart = null; 
            EditPoint epEnd = null; 
            epStart = elt.StartPoint.CreateEditPoint(); 
            // Do this because we move it later.
            epEnd = elt.EndPoint.CreateEditPoint(); 
            epStart.EndOfLine(); 
            if ( ( ( elt.IsCodeType ) & ( elt.Kind !=
              vsCMElement.vsCMElementDelegate ) ) ) 
            {
                builder.buildername = elt.Name + "Builder";
                builder.returnType = elt.Name;
                builder.rtVar = "p_" + elt.Name.ToLower();
                builder.body += @"public class " + elt.Name + @"Builder {
    private " + elt.Name + " " + builder.rtVar + " = new " + elt.Name + @"();
        ";
                //MessageBox.Show( "got type but not a delegate, named : " + elt.Name); 
                CodeType ct = null; 
                ct = ( ( EnvDTE.CodeType )( elt ) ); 
                CodeElements mems = null; 
                mems = ct.Members; 
                int i = 0; 
                for ( i=1; i<=ct.Members.Count; i++ ) 
                { 
                    CollapseElt( mems.Item( i ), mems, i , builder); 
                }
                builder.body += @"      public " + builder.returnType + @" build()
            {
                return " + builder.rtVar + @";
            }
                                    ";
                builder.body += "       }\r\n\r\n";
            } 
            else if ( ( elt.Kind == vsCMElement.vsCMElementNamespace ) ) 
            {
                builder.ns = elt.Name;
                builder.body += "namespace " + elt.Name + " {\r\n";
                //MessageBox.Show( "got a namespace, named: " + elt.Name); 
                CodeNamespace cns = null; 
                cns = ( ( EnvDTE.CodeNamespace )( elt ) ); 
                //MessageBox.Show( "set cns = elt, named: " + cns.Name); 

                CodeElements mems_vb = null; 
                mems_vb = cns.Members; 
                //MessageBox.Show( "got cns.members"); 
                int i = 0; 

                for ( i=1; i<=cns.Members.Count; i++ ) 
                { 
                    CollapseElt( mems_vb.Item( i ), mems_vb, i, builder ); 
                }
                builder.body += "}";
            } 
            else if (elt.Kind == vsCMElement.vsCMElementProperty)
            {
                var prop = elt as CodeProperty;
                //MessageBox.Show("got property named: " + prop.Type.AsString + " " + elt.Name);
                builder.body += @"  public " + builder.buildername + " " + elt.Name + " (" + prop.Type.AsString + " " + elt.Name.ToLower() + @")
        {
            " + builder.rtVar + "." + elt.Name + " = " + elt.Name.ToLower() + @";
            return this;
        }
                                    
        ";
            }

            return builder;
        }

        private void AddBuilderToSolution(Project proj, ProjectItem selectedProjectItem, string itemName, string projectPath, IVsHierarchy heirarchy)
        {
            string itemPath = Path.Combine(projectPath, itemName);
            if (!File.Exists(itemPath)) return;

            uint removeFileId;
            heirarchy.ParseCanonicalName(itemPath, out removeFileId);
            if (removeFileId < uint.MaxValue)
            {
                var itemToRemove = GetProjectItemFromHierarchy(heirarchy, removeFileId);
                if (itemToRemove!=null) itemToRemove.Remove();
            }

            // and add it to the project
            var addedItem = proj.ProjectItems.AddFromFile(itemPath); // selectedProjectItem.ProjectItems.AddFromFile(itemPath);
            addedItem.Properties.Item("ItemType").Value = "Compile";
        }

        private void AddBuilderFile(string content, string itemName, string projectPath, IVsHierarchy heirarchy)
        {
            string itemPath = Path.Combine(projectPath, itemName);
            if (!File.Exists(itemPath))
            {
                // create the new file
                using (var writer = new StreamWriter(itemPath))
                {
                    writer.Write(content);
                }
            }
        }

        private ProjectItem GetProjectItemFromHierarchy(IVsHierarchy pHierarchy, uint itemID)
        {
            object propertyValue;
            ErrorHandler.ThrowOnFailure(pHierarchy.GetProperty(itemID, (int)__VSHPROPID.VSHPROPID_ExtObject, out propertyValue));
            
            var projectItem = propertyValue as ProjectItem;
            if (projectItem == null) return null;
            
            return projectItem;
        }
        
        public static string GetErrorInfo()
        {
            string errText = null;
            var uiShell = (IVsUIShell)Package.GetGlobalService(typeof(IVsUIShell));
            if (uiShell != null)
            {
                uiShell.GetErrorInfo(out errText);
            }
            if (errText == null) return string.Empty;
            return errText;
        }

        private bool IsItemBuilderItem(IVsProject vsProject, uint itemid)
        {
            var buildPropertyStorage = vsProject as IVsBuildPropertyStorage;
            if (buildPropertyStorage == null) return false;

            bool isItemBuilderFile = false;

            string value;
            buildPropertyStorage.GetItemAttribute(itemid, IsBuilderFile, out value);
            if (string.Compare("true", value, true) == 0) isItemBuilderFile = true;

            // we need to special case web.config transform files
            if (!isItemBuilderFile)
            {
                //string pattern = @".*cs";
                string filepath;
                buildPropertyStorage.GetItemAttribute(itemid, "FullPath", out filepath);
                if (!string.IsNullOrEmpty(filepath))
                {
                    var fi = new System.IO.FileInfo(filepath);
                    //var regex = new System.Text.RegularExpressions.Regex(
                    //    pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    if (fi.Name.EndsWith(".cs"))
                    {
                        isItemBuilderFile = true;
                    }
                }
            }
            return isItemBuilderFile;
        }

        private bool ItemSupportsBuilders(IVsProject project, uint itemid)
        {
            string itemFullPath = null;

            if (ErrorHandler.Failed(project.GetMkDocument(itemid, out itemFullPath))) return false;

            // make sure its not a transform file itsle
            bool IsBuilderFile = IsItemBuilderItem(project, itemid);

            var transformFileInfo = new FileInfo(itemFullPath);
            bool isCSFile = transformFileInfo.Name.EndsWith(".cs");

            return (isCSFile && IsBuilderFile);
        }

        List<string> SupportedProjectExtensions = new List<string>
        {
            ".csproj",
            ".vbproj",
            ".fsproj"
        };

        private bool ProjectSupportsBuilders(IVsProject project)
        {
            string projectFullPath = null;
            if (ErrorHandler.Failed(project.GetMkDocument(VSConstants.VSITEMID_ROOT, out projectFullPath))) return false;

            string projectExtension = Path.GetExtension(projectFullPath);

            foreach (string supportedExtension in SupportedProjectExtensions)
            {
                if (projectExtension.Equals(supportedExtension, StringComparison.InvariantCultureIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        public static bool IsSingleProjectItemSelection(out IVsHierarchy hierarchy, out uint itemid)
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
    }

    public class myBuilder
    {
        public string ns { get; set; }
        public string buildername { get; set; }
        public string returnType { get; set; }
        public string rtVar { get; set; }
        public string body;
    }

    
}
