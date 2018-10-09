using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace VuePack
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", Vsix.Version, IconResourceID = 400)] // Info on this package for Help/About
    [ProvideAutoLoad(UIContextGuids80.SolutionHasSingleProject, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(UIContextGuids80.SolutionHasMultipleProjects, PackageAutoLoadFlags.BackgroundLoad)]
    [Guid(VuePackage.PackageGuidString)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class VuePackage : AsyncPackage
    {
        /// <summary>
        /// VuePackage GUID string.
        /// </summary>
        public const string PackageGuidString = "199cd44d-d4c2-4ec7-9b9d-4ce95c942d7b";
        private EnvDTE.SolutionEvents _events;

        private EnvDTE.DTE _dte;
        private EnvDTE.Events _dteEvents;
        private EnvDTE.DocumentEvents _documentEvents;

        private System.Collections.Generic.IEnumerable<EnvDTE.Project> _projects;

        /// <summary>
        /// Initializes a new instance of the <see cref="VuePackage"/> class.
        /// </summary>
        public VuePackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async System.Threading.Tasks.Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            EnvDTE80.DTE2 dte2 = await GetServiceAsync(typeof(EnvDTE.DTE)) as EnvDTE80.DTE2;
            Microsoft.Assumes.Present(dte2);

            _events = dte2.Events.SolutionEvents;
            _events.AfterClosing += delegate { DirectivesCache.Clear(); };

            _dte = await GetServiceAsync(typeof(SDTE)) as EnvDTE.DTE;
            Microsoft.Assumes.Present(_dte);
            _dteEvents = _dte.Events;
            _documentEvents = _dteEvents.DocumentEvents;
            _documentEvents.DocumentSaved += OnDocumentSaved;

            IVsSolution solution = await GetServiceAsync(typeof(IVsSolution)) as IVsSolution;
            _projects = GetProjects(solution);
            
            await ParseVueTemplate.InitializeAsync(this);
        }

        #endregion


        private void OnDocumentSaved(EnvDTE.Document document)
        {
            string path = document.FullName;
            string ext = System.IO.Path.GetExtension(path);
            if (ext == null || ext.ToLowerInvariant() != ".vue") return;

            /* When saving a .VUE file, search project directory for
             * matching .TS file. Open TS file and find/replace text 
             * between two multi-line string ticks (``) with next text
             * from Vue <template>
             */

            string Name = document.Name.Replace(".vue", ".ts");

            foreach (var project in _projects)
            {
                if (!string.IsNullOrEmpty(project.FullName))
                {
                    string startPath = System.IO.Path.GetDirectoryName(project.FullName);

                    //search within project for matching TS file
                    foreach (string TSfile in System.IO.Directory.GetFiles(startPath, Name, System.IO.SearchOption.AllDirectories)) //"*.ts"
                    {
                        UpdateFile(TSfile, GetDocumentText(document));
                    }
                }
            }
        }

        private static string GetDocumentText(EnvDTE.Document document)
        {
            EnvDTE.TextDocument textDocument = (EnvDTE.TextDocument)document.Object("TextDocument");
            EnvDTE.EditPoint editPoint = textDocument.StartPoint.CreateEditPoint();
            string context = editPoint.GetText(textDocument.EndPoint);
            context = context.Replace(System.Environment.NewLine, "");
            context = context.Replace("\n", "");
            context = System.Text.RegularExpressions.Regex.Replace(context, " {2,}", " "); //REmoves all whitespaces
            //context = context.Replace("<template>", "").Replace("</template>", "");
            return context.Trim();
        }

        private static void UpdateFile(string filePath, string vueTemplate = "")
        {
            if (string.IsNullOrEmpty(filePath))
            {
                //System.Console.WriteLine("ERROR: Must have a full file path.");
                return;
            }

            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.Append("`").Append(vueTemplate).Append("`");

            if (System.IO.File.Exists(filePath))
            {
                string text = System.IO.File.ReadAllText(filePath);

                int start = text.IndexOf("`") + 1;
                int end = text.LastIndexOf("`,");

                if (end > start)
                {
                    int cnt = end - start;

                    text = text.Remove(start, cnt);
                    text = text.Replace("``", sb.ToString());

                    System.IO.File.WriteAllText(filePath, text);
                }
            }
        }

        public static System.Collections.Generic.IEnumerable<EnvDTE.Project> GetProjects(IVsSolution solution)
        {
            foreach (IVsHierarchy hier in GetProjectsInSolution(solution))
            {
                EnvDTE.Project project = GetDTEProject(hier);
                if (project != null)
                    yield return project;
            }
        }

        public static System.Collections.Generic.IEnumerable<IVsHierarchy> GetProjectsInSolution(IVsSolution solution)
        {
            return GetProjectsInSolution(solution, __VSENUMPROJFLAGS.EPF_LOADEDINSOLUTION);
        }

        public static System.Collections.Generic.IEnumerable<IVsHierarchy> GetProjectsInSolution(IVsSolution solution, __VSENUMPROJFLAGS flags)
        {
            if (solution == null)
                yield break;

            IEnumHierarchies enumHierarchies;
            Guid guid = Guid.Empty;
            solution.GetProjectEnum((uint)flags, ref guid, out enumHierarchies);
            if (enumHierarchies == null)
                yield break;

            IVsHierarchy[] hierarchy = new IVsHierarchy[1];
            uint fetched;
            while (enumHierarchies.Next(1, hierarchy, out fetched) == Microsoft.VisualStudio.VSConstants.S_OK && fetched == 1)
            {
                if (hierarchy.Length > 0 && hierarchy[0] != null)
                    yield return hierarchy[0];
            }
        }

        public static EnvDTE.Project GetDTEProject(IVsHierarchy hierarchy)
        {
            if (hierarchy == null)
                throw new ArgumentNullException("hierarchy");

            //var u = Microsoft.VisualStudio.VSConstants.VSITEMID.Root;//4294967294 VSConstants.VSITEMID_ROOT
            object obj;
            hierarchy.GetProperty(4294967294, (int)__VSHPROPID.VSHPROPID_ExtObject, out obj);
            return obj as EnvDTE.Project;
        }

    }
}
