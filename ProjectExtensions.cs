using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace EfModelGenerator
{
    public static class ProjectExtensions
    {

        public static bool IsWebProject(this Project project)
        {
            return project.GetProjectTypes().Any<string>((Func<string, bool>)(g =>
            {
                if (!g.EqualsIgnoreCase("{349C5851-65DF-11DA-9384-00065B846F21}"))
                    return g.EqualsIgnoreCase("{E24C65DC-7377-472B-9ABA-BC803B73C61A}");
                return true;
            }));
        }

        public static string ProjectPath(this Project project)
        {
            return project.Properties.Item("FullPath").Value as string;
        }

        /// <summary>
        /// Get the executing projects configuration
        /// </summary>
        /// <returns></returns>
        public static System.Configuration.Configuration GetProjectConfiguration(this Project project)
        {
            XDocument document = XDocument.Load(Path.Combine(project.ProjectPath() as string, project.IsWebProject() ? "Web.config" : "App.config"));

            string tempFileName = Path.GetTempFileName();
            document.Save(tempFileName);
            return System.Configuration.ConfigurationManager.OpenMappedExeConfiguration(new ExeConfigurationFileMap()
            {
                ExeConfigFilename = tempFileName
            }, ConfigurationUserLevel.None);
        }

        private static IEnumerable<string> GetProjectTypes(this Project project)
        {
            IVsSolution service;
            using (ServiceProvider serviceProvider = new ServiceProvider((Microsoft.VisualStudio.OLE.Interop.IServiceProvider)project.DTE))
                service = (IVsSolution)serviceProvider.GetService(typeof(IVsSolution));
            IVsHierarchy ppHierarchy;
            int projectOfUniqueName = service.GetProjectOfUniqueName(project.UniqueName, out ppHierarchy);
            if (projectOfUniqueName != 0)
                Marshal.ThrowExceptionForHR(projectOfUniqueName);
            string pbstrProjTypeGuids;
            int projectTypeGuids = ((IVsAggregatableProject)ppHierarchy).GetAggregateProjectTypeGuids(out pbstrProjTypeGuids);
            if (projectTypeGuids != 0)
                Marshal.ThrowExceptionForHR(projectTypeGuids);
            return (IEnumerable<string>)pbstrProjTypeGuids.Split(';');
        }

        public static bool EqualsIgnoreCase(this string s1, string s2)
        {
            return string.Equals(s1, s2, StringComparison.OrdinalIgnoreCase);
        }

        public static CodeElement GetNameSpace(this ProjectItem projectItem)
        {
            var model = projectItem.FileCodeModel;
            foreach (CodeElement element in model.CodeElements)
            {
                if (element.Kind == vsCMElement.vsCMElementNamespace)
                {
                    return element;
                }
            }
            return null;
        }

        public static bool IsDerivedFromDbContext(this ProjectItem item)
        {

            var model = item.FileCodeModel;

            if (model == null) return false;

            var isDerivedFromDbContext = false;

            foreach (CodeElement element in model.CodeElements)
            {

                if (element is EnvDTE.CodeNamespace)
                {
                    var namespaceElement = element as EnvDTE.CodeNamespace;

                    foreach (var property in namespaceElement.Members)
                    {
                        var codeType = property as CodeType;
                        if (codeType == null) continue;

                        foreach (var member in codeType.Bases)
                        {
                            var codeClass = member as CodeClass;

                            if (codeClass == null) continue;
                            var name = codeClass.Name;

                            if (!string.IsNullOrEmpty(name) && name.EqualsIgnoreCase("DbContext"))
                            {
                                isDerivedFromDbContext = true;
                                break;
                            }
                        }
                    }
                }
            }
            return isDerivedFromDbContext;
        }
    }
}
