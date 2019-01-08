using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace VuePack
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ParseVueTemplate
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("7657fbed-f092-4ce0-bb76-73a39ff14345");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;
        private static Microsoft.VisualStudio.TextManager.Interop.IVsTextManager2 _textManagerSvc;
        private static EnvDTE80.DTE2 _dte2;

        /// <summary>
        /// Initializes a new instance of the <see cref="ParseVueTemplate"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private ParseVueTemplate(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ParseVueTemplate Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in ParseVueTemplate's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            _textManagerSvc = await package.GetServiceAsync(typeof(Microsoft.VisualStudio.TextManager.Interop.SVsTextManager)) as Microsoft.VisualStudio.TextManager.Interop.IVsTextManager2;
            Microsoft.Assumes.Present(_textManagerSvc);

            _dte2 = await package.GetServiceAsync(typeof(EnvDTE.DTE)) as EnvDTE80.DTE2;
            Microsoft.Assumes.Present(_dte2);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new ParseVueTemplate(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            TextViewSelection selection = GetSelection();

            string Name = _dte2.ActiveDocument.Name;
            string fullpath = System.IO.Path.GetTempPath() + Name.Replace(".ts", ".vue");

            if (!_dte2.ItemOperations.IsFileOpen(fullpath))
            {
                var sw = System.IO.File.CreateText(fullpath);
                sw.Write(selection.Text.Replace("> <", $"> {Environment.NewLine} <"));
                sw.WriteLine("");
                sw.WriteLine($"<!--# sourceURL ={_dte2.ActiveDocument.FullName}-->");
                sw.Close();

                EnvDTE.Window window = _dte2.ItemOperations.OpenFile(fullpath);
                window.Activate();
                _dte2.ExecuteCommand("Edit.FormatDocument");
            }
            else
            {
                EnvDTE.Document document = _dte2.Documents.Item(fullpath);
                document.Activate();
            }
        }

        private TextViewSelection GetSelection()
        {
            Microsoft.VisualStudio.TextManager.Interop.IVsTextView view;
            int result = _textManagerSvc.GetActiveView2(1, null, (uint)Microsoft.VisualStudio.TextManager.Interop._VIEWFRAMETYPE.vftCodeWindow, out view);

            view.GetSelection(out int startLine, out int startColumn, out int endLine, out int endColumn);//end could be before beginning
            var start = new TextViewPosition(startLine, startColumn);
            var end = new TextViewPosition(endLine, endColumn);

            view.GetSelectedText(out string selectedText);

            view.GetCaretPos(out int line, out int col);


            TextViewSelection selection = new TextViewSelection(start, end, selectedText);
            return selection;
        }


    }
}
