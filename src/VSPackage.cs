using System;
using System.Runtime.InteropServices;
using System.Threading;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Web.Editor.Controller.Constants;

namespace VuePack
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", Vsix.Version, IconResourceID = 400)]
    [ProvideAutoLoad(UIContextGuids80.SolutionHasSingleProject, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(UIContextGuids80.SolutionHasMultipleProjects, PackageAutoLoadFlags.BackgroundLoad)]
    [Guid("b2295a37-9de5-4be8-8a5e-7bbb7ecdc3ca")]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public sealed class VuePackage : AsyncPackage
    {
        private SolutionEvents _events;

        private DTE _dte;
        private Events _dteEvents;
        private DocumentEvents _documentEvents;

        private System.Collections.Generic.IEnumerable<Project> _projects;

        protected override async System.Threading.Tasks.Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            var dte = await GetServiceAsync(typeof(DTE)) as DTE2;

            _events = dte.Events.SolutionEvents;
            _events.AfterClosing += delegate { DirectivesCache.Clear(); };

            _dte = await GetServiceAsync(typeof(SDTE)) as DTE;
            _dteEvents = _dte.Events;
            _documentEvents = _dteEvents.DocumentEvents;
            _documentEvents.DocumentSaved += OnDocumentSaved;

            var solution = await GetServiceAsync(typeof(IVsSolution)) as IVsSolution;            
            _projects = GetProjects(solution);
            //System.Diagnostics.Debug.Print("");
        }

        private void OnDocumentSaved(Document document)
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
                bool processing = true;
                string startPath = System.IO.Path.GetDirectoryName(project.FullName);

                //search each project to find original file
                foreach (string file in System.IO.Directory.GetFiles(startPath, document.Name, System.IO.SearchOption.AllDirectories))
                {
                    //search within project for related TS file
                    foreach (string TSfile in System.IO.Directory.GetFiles(startPath, "*.ts", System.IO.SearchOption.AllDirectories))
                    {
                        if (TSfile.IndexOf(Name) > -1)
                        {
                            UpdateFile(TSfile, GetDocumentText(document));
                            processing = false;
                            break;
                        }

                        if (!processing)
                            break;
                    }

                }
            }
        }

        private static string GetDocumentText(Document document)
        {
            TextDocument textDocument = (TextDocument)document.Object("TextDocument");
            EditPoint editPoint = textDocument.StartPoint.CreateEditPoint();
            string context = editPoint.GetText(textDocument.EndPoint);
            context = context.Replace(System.Environment.NewLine, "");
            context = System.Text.RegularExpressions.Regex.Replace(context, " {2,}", " "); //REmoves all whitespaces
            context = context.Replace("<template>", "").Replace("</template>", "");
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
                int cnt = text.LastIndexOf("`") - start;

                text = text.Remove(start, cnt);
                text = text.Replace("``", sb.ToString());

                System.IO.File.WriteAllText(filePath, text);
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
            while (enumHierarchies.Next(1, hierarchy, out fetched) == VSConstants.S_OK && fetched == 1)
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
