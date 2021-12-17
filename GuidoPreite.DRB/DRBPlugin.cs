using System;
using System.ComponentModel.Composition;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Reflection;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace GuidoPreite.DRB
{
    [Export(typeof(IXrmToolBoxPlugin)),
        ExportMetadata("Name", "Dataverse REST Builder"),
        ExportMetadata("Description", "Create and Execute requests against the Dataverse Web API endpoint"),
        ExportMetadata("SmallImageBase64", "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsIAAA7CARUoSoAAAAKuSURBVFhH7VZNaBNREJ55uxuTlKoHbf1BPPSggj+FetKDNkGI/bN4EEXxqog59aBQaU9SQYqgoIiHeBEkCoI2WgVJkSo9qVih9aI3RfsjKDVm8/aNs29f09S22lNzaL7Dvp353nsz883ALlRQwbIHWsmOm+a9LEAnM0DmvSyYSYDglF6XCgha+WICheb96K9Lhem4QltlRNkTmLcFdt9AJ6e2WxsEOSB8LZW6DW2N45p/nE3y0VjAkwuKXknrdwqamn5YmecxJCupS1M0xfywdMK3ILFnUu83+E8L6DAqinHwCb6gCoF6bIFvIs8GN/gsetSEBAkk+szrakS8YqvIDc2BVY8I7XzOD8AU9jgyf9fn5sOCLSDEbzJKZ2VL4yEE1c6Xri+4stfQvjJTBRc6C7/GWjnMV47VAOn3IcP6oXv47HFAMcTWlsA5F4uaAbclluFyhrm8hHHpUh2H6kKRtW1srWL7HRzZ7gYk0wpa7b7sec60gcW4btxzMG8CLJvuTylY7jEkFTamj5Vc3QsSeI8DjEgrdDpwK/3kC45xUif5RYEQ9ZDKlp4tYlEKQDptEeA2bssX4+EI9F2QtYOD30HAnbYsbNX+YvLihKyt3kVAXTzdR0M1JeqV4F8JCMhZ1eEH/ZucqjWXeBA3si8VUAxEmRfOuAC8yNl4XPk56O6euU95Dnz8uYI3RoxHmnUWFkyAs97sKG/Uc8IjvK2Dq7wvo3DZ0L4C/JgEt3nfKKvzkvfHI3sPrAtIPm/BQ6cKP/HwdrFKT90o9BtqFmyzzoIiusBK1nHviPs+gYLeugfjHwwNSoirgqgWampyvuTiyeAZRTJObt6xbPHI85QeRp6lHHERsjk2xEYwHH+h8i0oKuBd613Sz/H0j9DM/0CZgHYmW9ZfsgoqWO4A+APE3Qef6LHYmAAAAABJRU5ErkJggg=="),
        ExportMetadata("BigImageBase64", "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"),
        ExportMetadata("BackgroundColor", "White"),
        ExportMetadata("PrimaryFontColor", "Black"),
        ExportMetadata("SecondaryFontColor", "Gray")]

    public class DRBPlugin : PluginBase
    {
        public override IXrmToolBoxPluginControl GetControl()
        {
            return new DRBPluginControl();
        }

        /// <summary>
        /// Constructor 
        /// </summary>
        public DRBPlugin()
        {
            // If you have external assemblies that you need to load, uncomment the following to 
            // hook into the event that will fire when an Assembly fails to resolve
            // AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(AssemblyResolveEventHandler);
        }

        /// <summary>
        /// Event fired by CLR when an assembly reference fails to load
        /// Assumes that related assemblies will be loaded from a subfolder named the same as the Plugin
        /// For example, a folder named Sample.XrmToolBox.MyPlugin 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        private Assembly AssemblyResolveEventHandler(object sender, ResolveEventArgs args)
        {
            Assembly loadAssembly = null;
            Assembly currAssembly = Assembly.GetExecutingAssembly();

            // base name of the assembly that failed to resolve
            var argName = args.Name.Substring(0, args.Name.IndexOf(","));

            // check to see if the failing assembly is one that we reference.
            List<AssemblyName> refAssemblies = currAssembly.GetReferencedAssemblies().ToList();
            var refAssembly = refAssemblies.Where(a => a.Name == argName).FirstOrDefault();

            // if the current unresolved assembly is referenced by our plugin, attempt to load
            if (refAssembly != null)
            {
                // load from the path to this plugin assembly, not host executable
                string dir = Path.GetDirectoryName(currAssembly.Location).ToLower();
                string folder = Path.GetFileNameWithoutExtension(currAssembly.Location);
                dir = Path.Combine(dir, folder);

                var assmbPath = Path.Combine(dir, $"{argName}.dll");

                if (File.Exists(assmbPath))
                {
                    loadAssembly = Assembly.LoadFrom(assmbPath);
                }
                else
                {
                    throw new FileNotFoundException($"Unable to locate dependency: {assmbPath}");
                }
            }

            return loadAssembly;
        }
    }
}