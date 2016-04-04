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

using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Editor;

using Microsoft.VisualStudio.Text;

using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;

//using AsmLanguage;

namespace MenuCommandUncomment
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
    /// 
   
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidMenuCommandUncommentPkgString)]
    public sealed class MenuCommandUncommentPackage : Package
    {
        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public MenuCommandUncommentPackage()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));
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
            Debug.WriteLine (string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
            base.Initialize();

            // Add our command handlers for menu (commands must exist in the .vsct file)
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if ( null != mcs )
            {
                // Create the command for the menu item.
                CommandID menuCommandID = new CommandID(GuidList.guidMenuCommandUncommentCmdSet, (int)PkgCmdIDList.cmdidUnComment);
                MenuCommand menuItem = new MenuCommand(MenuItemCallback, menuCommandID );
                mcs.AddCommand( menuItem );
            }
        }
        #endregion

        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            IVsTextManager txtMgr = (IVsTextManager)GetService(typeof(SVsTextManager));
            IVsTextView vTextView = null;
            int mustHaveFocus = 1;
            txtMgr.GetActiveView(mustHaveFocus, null, out vTextView);

            IVsUserData userData = vTextView as IVsUserData;
            if (userData == null)
            {
                Console.WriteLine("No text view is currently open");
                return;
            }
            IWpfTextViewHost viewHost;
            object holder;
            Guid guidViewHost = DefGuidList.guidIWpfTextViewHost;
            userData.GetData(ref guidViewHost, out holder);
            viewHost = (IWpfTextViewHost)holder;

           // Connector.Execute(viewHost);    
            IWpfTextView view = viewHost.TextView;

            string filename = "";
                
            ITextDocument document;
            view.TextDataModel.DocumentBuffer.Properties.TryGetProperty(typeof(Microsoft.VisualStudio.Text.ITextDocument), out document);
            if ((document != null) && (document.TextBuffer != null))
                filename = document.FilePath;
            string ext = System.IO.Path.GetExtension(filename).ToLower();
            string comment = "";
            switch (ext)
            {
                case ".asm":
                case ".inc":
                case ".ini":
                    comment = ";";
                    break;

                case ".vb":
                    comment = "'";
                    break;

                case ".sql":
                    comment = "--";
                    break;

                case ".c":
                case ".cc":
                case ".d":
                case ".cpp":
                case ".cxx":
                case ".c++":
                case ".h":
                case ".hpp":
                case ".hxx":
                case ".cs":
                    comment = "//";
                    break;

                case ".txt":
                    comment = "#";
                    break;
            }
            if (comment == "") return;

            Regex rx = new Regex(@"^\s*" + comment, RegexOptions.Compiled);
            Regex rxtabs = new Regex(@"^(?'tabs'\s*).*", RegexOptions.Compiled);
            //Add a comment on the selected text. 
            if (!view.Selection.IsEmpty)
            {
                var span = view.Selection.SelectedSpans[0];

                string s = span.Snapshot.GetText(span);
                StringReader sr = new StringReader(s);
                StringBuilder sb = new StringBuilder();

                int commented_lines = 0;
                List<string> strings = new List<string>();

                string textline = null;
                while ((textline = sr.ReadLine()) != null)
                {
                    strings.Add(textline);
                    if (rx.Match(textline).Success) commented_lines++;
                }
                sr.Close();

                bool first = true;
                if (2 * commented_lines > strings.Count)
                {
                    // uncomment
                    foreach (string line in strings)
                    {
                        if (first)
                            first = false;
                        else
                            sb.AppendLine();

                        if (rx.Match(line).Success)
                        {
                            int pos1 = line.IndexOf(comment);
                            sb.Append(line.Remove(pos1, comment.Length));
                        }
                        else sb.Append(line);
                    }
                }
                else
                {
                    // comment
                    foreach (string line in strings)
                    {
                        if (first)
                            first = false;
                        else
                            sb.AppendLine();
                        var m = rxtabs.Match(line);
                        string tabs = m.Groups["tabs"].Value;
                        string txt = line.Substring(tabs.Length, line.Length - tabs.Length);
                        sb.Append(tabs);
                        if (txt == "") continue;
                        sb.Append(comment);
                        sb.Append(txt);
                    }
                }

                var edit = view.TextBuffer.CreateEdit();
                if (s.EndsWith("\r\n")) sb.AppendLine();
                SnapshotPoint sp = view.Selection.SelectedSpans[0].Start;
                edit.Replace(view.Selection.SelectedSpans[0], sb.ToString());
                edit.Apply();
                view.Selection.Select(new SnapshotSpan(view.TextBuffer.CurrentSnapshot, sp.Position, sb.Length), false);
            }
            //
        }

    }
}
