using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using System.Text;
using System.Text.RegularExpressions;
using EnvDTE;
using Microsoft.VisualBasic;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;

namespace SoftwareNinjas.TestOriented.VisualStudio
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
    // This attribute is used to register the information needed to show the this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.PkgString)]
    public sealed class Package : Microsoft.VisualStudio.Shell.Package
    {
        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public Package()
        {
            Trace.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this));
        }



        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation
        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            Trace.WriteLine (string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this));
            base.Initialize();

            // Add our command handlers for menu (commands must exist in the .vsct file)
            var mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if ( null != mcs )
            {
                // Create the command for the menu item.
                var menuCommandId = new CommandID(GuidList.CmdSet, (int)PkgCmdIDList.cmdidGenerateTestStub);
                var menuItem = new MenuCommand(GenerateTestStub, menuCommandId );
                mcs.AddCommand( menuItem );
            }
        }
        #endregion

        private const string CrLf = "\r\n";

        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void GenerateTestStub(object sender, EventArgs e)
        {
            var uiShell = (IVsUIShell)GetService(typeof(SVsUIShell));
            var clsid = Guid.Empty;

            var dte = (DTE) GetGlobalService(typeof (DTE));
            var document = dte.ActiveDocument;
            var sel = (TextSelection) document.Selection;
            var pnt = sel.ActivePoint;
            var projectItem = document.ProjectItem;
            var fileCodeModel = projectItem.FileCodeModel;
            var cut = (CodeClass) fileCodeModel.CodeElementFromPoint(pnt, vsCMElement.vsCMElementClass);
            // TODO: stop right here if the class already ends in "Test"
            CodeFunction mut = null;
            try
            {
                mut = (CodeFunction) fileCodeModel.CodeElementFromPoint(pnt, vsCMElement.vsCMElementFunction);
            }
            catch (COMException)
            {
                // there may not be any methods, or cursor was not near any method
            }
            // TODO: stop right here if the method is not visible to the test class

            var testClassName = cut.Name + "Test";
            var testClassFileName = testClassName + ".cs";
            var testClassFile = dte.Solution.FindProjectItem(testClassFileName);
            if (testClassFile == null)
            {
                var put = projectItem.ContainingProject;
                var putName = put.Name;
                var candidates = FindTestProject(putName, dte);
                Project testProject;
                if (candidates.Count == 0)
                {
                    int result;
                    ErrorHandler.ThrowOnFailure(
                        uiShell.ShowMessageBox(
                            0,
                            ref clsid,
                            "Sucks to be you",
                            string.Format(CultureInfo.CurrentCulture, 
                                "There exists no test project for the project '{0}'.", putName),
                            string.Empty,
                            0,
                            OLEMSGBUTTON.OLEMSGBUTTON_OK,
                            OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
                            OLEMSGICON.OLEMSGICON_INFO,
                            0,        // false
                            out result
                        )
                    );
                    return;
                }
                if (candidates.Count == 1)
                {
                    testProject = candidates[0];
                }
                else
                {
                    var sb = new StringBuilder();
                    var c = 0;
                    foreach (var project in candidates)
                    {
                        sb.Append(c);
                        sb.Append(" = ");
                        sb.AppendLine(project.Name);
                        c = c + 1;
                    }
                    var selectionString = 
                        Interaction.InputBox(sb.ToString(), "Which project should the test class be added to?", "0");
                    var selectedIndex = Convert.ToInt32(selectionString, 10);
                    testProject = candidates[selectedIndex];
                }
                var testProjectRoot = Path.GetDirectoryName(testProject.FullName);
                var testClassFolder = testProjectRoot;
                var targetNamespace = testProject.Name;
                if (cut.Namespace.Name.StartsWith(put.Name + "."))
                {
                    var rootOffset = cut.Namespace.Name.Remove(0, put.Name.Length + 1);
                    testClassFolder = Path.Combine(testClassFolder, rootOffset);
                    targetNamespace = targetNamespace + "." + rootOffset;
                }
                var absolutePathToTestClass = Path.Combine(testClassFolder, testClassFileName);
                var testClassContents = new StringBuilder();
                testClassContents.AppendLine("using System;");
                testClassContents.AppendLine("using Microsoft.VisualStudio.TestTools.UnitTesting;");

                testClassContents.AppendFormat("using {0};", cut.Namespace.Name).AppendLine();
                testClassContents.AppendLine();

                testClassContents.AppendFormat("namespace {0}", targetNamespace).AppendLine();
                testClassContents.AppendLine("{");
                testClassContents.AppendLine("    /// <summary>");
                testClassContents.AppendFormat(@"    /// A class to test <see cref=""{0}""/>.", cut.Name).AppendLine();
                testClassContents.AppendLine("    /// </summary>");
                testClassContents.AppendLine("    [TestClass]");
                testClassContents.AppendFormat("    public class {0}", testClassName).AppendLine();
                testClassContents.AppendLine("    {");
                testClassContents.AppendLine("    }");
                testClassContents.AppendLine("}");
                File.WriteAllText(absolutePathToTestClass, testClassContents.ToString());

                testClassFile = testProject.ProjectItems.AddFromFile(absolutePathToTestClass);
            }
            testClassFile.Open(EnvDTE.Constants.vsViewKindCode);
            testClassFile.Document.Activate();

            var testFileElements = testClassFile.FileCodeModel.CodeElements;
            var testClass = (CodeClass) FindElement(testFileElements, testClassName, vsCMElement.vsCMElementClass);

            if (mut != null)
            {
                try
                {
                    dte.UndoContext.Open("Insert test method");

                    CodeFunction testMethod = testClass.AddFunction(mut.Name + "TODO", vsCMFunction.vsCMFunctionFunction,
                                                                    vsCMTypeRef.vsCMTypeRefVoid);

                    var editPoint = testMethod.StartPoint.CreateEditPoint();

                    // Adding the "[TestMethod]" attribute this way avoid the DTE adding it as "[Test()]" when using AddAttribute()
                    editPoint.Insert("[TestMethod]");
                    editPoint.Insert(CrLf);
                    editPoint.Indent(null, 2);

                    editPoint.LineDown(2);
                    editPoint.StartOfLine();
                    // TODO: generate local variables to match the parameter names & types
                    //var parameters = mut.Parameters;
                    //foreach (var codeElement2 in parameters)
                    //{
                    //    CodeParameter = codeElement2;
                    //    editPoint.Indent(CodeParameter.Name);
                    //}

                    // TODO: check if the methodUnderTest is an instance or a static method
                    // TODO: check if the methodUnderTest returns a value or not
                    // TODO: use editPoint.Insert("Text") instead of the following?
                    editPoint.Indent(null, 3);
                    editPoint.Insert("var actual = " + cut.Name + "." + mut.Name + "();");
                    editPoint.Insert(CrLf);
                    editPoint.Indent(null, 3);
                    editPoint.Insert("Assert.AreEqual(0, actual);");

                    // place the cursor at the method call
                    editPoint.LineUp(1);
                    var selection = (TextSelection) testClassFile.Document.Selection;
                    selection.GotoLine(editPoint.Line);
                    selection.EndOfLine(false);
                    selection.CharLeft(false, 2);
                }
                finally
                {
                    dte.UndoContext.Close();
                }
            }
        }

        internal List<Project> FindTestProject(string projectUnderTestName, DTE dte)
        {
            var candidates = new List<Project>(dte.Solution.Projects.Count);

            foreach (var project in RecurseProjects(dte))
            {
                if (Regex.IsMatch(project.Name, projectUnderTestName + @"\..*Test"))
                {
                    candidates.Add(project);
                }
            }

            return candidates;
        }

        internal List<Project> RecurseProjects(DTE dte)
        {
            var result = new List<Project>(dte.Solution.Projects.Count);

            foreach (Project project in dte.Solution.Projects)
            {
                InnerRecurseProjects(project, result);
            }

            return result;
        }

        private const string SolutionFolder = "{66A26720-8FB5-11D2-AA7E-00C04F688DDE}";
        internal void InnerRecurseProjects(Project project, List<Project> projects)
        {
            
            if (SolutionFolder == project.Kind)
            {
                foreach (ProjectItem projectItem in project.ProjectItems)
                {
                    if (projectItem.SubProject != null)
                    {
                        InnerRecurseProjects(projectItem.SubProject, projects);
                    }
                }
            }
            else
            {
                projects.Add(project);
            }
        }

        internal CodeElement FindElement(CodeElements elements, string name, vsCMElement kind)
        {
            foreach (CodeElement element in elements)
            {
                if (element.Kind == kind)
                {
                    if (element.Name == name)
                    {
                        return element;
                    }
                }
                var children = element.Children;
                if (children != null && children.Count > 0)
                {
                    var result = FindElement(children, name, kind);
                    if (result != null)
                    {
                        return result;
                    }
                }
            }
            return null;
        }
    }
}
