using Microsoft.Win32;
using System;
using System.IO;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Xml;

namespace WebAddInSideloader
{
    class Program
    {
        private const string REG_KEY = @"Software\Microsoft\Office\16.0\WEF\Developer";
        private const string VALUE_UseDirectDebugger = "UseDirectDebugger";
        private const string VALUE_UseLiveReload = "UseLiveReload";
        private const string VALUE_UseWebDebugger = "UseWebDebugger";
        private const string PATH_IdTag = "//oa:OfficeApp/oa:Id";
        private const string XML_NS = "oa";
        private const string XML_NS_PATH = "http://schemas.microsoft.com/office/appforoffice/1.1";
        /// <summary>
        /// Main entry point - For install we have to write this:
        /// Computer\HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\Developer
        /// PATH_TO_XML = PATH_TO_XML
        ///Computer\HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\Developer\{GUID}
        /// UseDirectDebugger = 1
        /// UseLiveReload = 1
        /// UseWebDebugger = 1
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] PobjArgs)
        {
            bool LbolSuccess = false;
            // process the arguments
            ProcessedArguments LobjArgs = new ProcessedArguments(PobjArgs);
            if (LobjArgs.Help)
            {
                Console.WriteLine(HelpString());
                LbolSuccess = true;
            }
            else if (LobjArgs.Processed)
            {
                if (LobjArgs.Install) LbolSuccess = Install(LobjArgs.ManifestPath, LobjArgs.InstallPath);
                if (LobjArgs.Uninstall) LbolSuccess = Uninstall(LobjArgs.InstalledManifestFullName);
                if (LobjArgs.Update) LbolSuccess = Update(LobjArgs.ManifestPath, LobjArgs.InstallPath);
            }

            // were we successful
            if(LbolSuccess)
            {
                Console.WriteLine("The process completed successfully!");
            }
            else
            {
                Console.WriteLine("The process DID NOT complete successfully.");
            }
#if DEBUG
            Console.ReadLine();
#endif
        }

        /// <summary>
        /// The Help file
        /// </summary>
        /// <returns></returns>
        private static string HelpString()
        {
            return "Web AddIn Sideloader Command Line Utility\n" +
                   "=========================================\n" +
                   "Version: " + Assembly.GetExecutingAssembly().GetName().Version + "\n\n" +
                   "This utlity is to allow enterprise organizations without Office 365 or " +
                   "centeralized add-in governance to be able to \ninstall web add-ins to users " +
                   "desktops, requiring no effort by the users to have the add-ins installed " +
                   "and available \nfor use.\n\n" +
                   "To use this utility, you will use only one of these required switches: \n\n" +
                   "\t-install\tInstalls the add-in\n" +
                   "\t-uninstall\tUninstalls the add-in\n" +
                   "\t-update\t\tUpdates the add-in\n\n" +
                   "You will also need to provide BOTH of these switches with any of the above options: \n\n" +
                   "\t-installPath [local folder path]\n" +
                   "\t-manifestPath [centralized manifest XML file]\n\n" +
                   "\tNOTE: The install path folder MUST exist. \n\n" +
                   "The following are some examples of usage: \n\n" +
                   "Set-WebAddin -install -installPath c:\\add-in -manifestPath \\\\server\\share\\manifest.xml\n" +
                   "Set-WebAddin -uninstall -installPath c:\\add-in -manifestPath \\\\server\\share\\manifest.xml\n" +
                   "Set-WebAddin -update -installPath c:\\add-in -manifestPath \\\\server\\share\\manifest.xml\n";
        }

        /// <summary>
        /// Installs the Web Add-in
        /// </summary>
        /// <param name="PstrManifestPath"></param>
        /// <param name="PstrInstallPath"></param>
        private static bool Install(string PstrManifestPath, string PstrInstallPath)
        {
            try
            {
                // copy the manifest down to the install path
                Console.WriteLine("Accessing the manifest file: " + PstrManifestPath);

                string localManifestFullPath = "";

                if (PstrManifestPath.ToLower().StartsWith("http"))
                {
                    localManifestFullPath = DownloadFileFromWeb(PstrManifestPath, PstrInstallPath);
                }
                else
                {
                    localManifestFullPath = DownloadFile(PstrManifestPath, PstrInstallPath);
                }

                Console.WriteLine("Copy complete. Analyzing file...");
                string LstrId = GetManifestId(localManifestFullPath);
                
                if (LstrId == null) throw new Exception("Invalid manifest file detected.");
                Console.WriteLine("Writing Registry entries...");
                RegistryKey LobjKey = Registry.CurrentUser.OpenSubKey(REG_KEY, true);
                if (LobjKey == null) LobjKey = Registry.CurrentUser.CreateSubKey(REG_KEY, true);
                if (LobjKey == null) throw new Exception("Unable to create or write to the registry path: HKCU\\" + REG_KEY);
                LobjKey.SetValue(localManifestFullPath, localManifestFullPath, RegistryValueKind.String);
                LobjKey = LobjKey.CreateSubKey(LstrId);
                LobjKey.SetValue(VALUE_UseDirectDebugger, 1, RegistryValueKind.DWord);
                LobjKey.SetValue(VALUE_UseLiveReload, 0, RegistryValueKind.DWord);
                LobjKey.SetValue(VALUE_UseWebDebugger, 0, RegistryValueKind.DWord);
                Console.WriteLine("Registry keys completed.");
                Console.WriteLine("Install process completed.");
                return true;
            }
            catch(Exception PobjEx)
            {
                PobjEx.HandleException(true, "Either the manifest path is invalid, cannot be accessed, or " +
                                             "the manifest is not a valid XML file, is not a Web Add-in " +
                                             "manifest, or the local path does not exist, or the local " +
                                             "file is still in use.");
                return false;
            }
        }

        /// <summary>
        /// Uninstalls the web add-in
        /// </summary>
        /// <param name="PstrLocalManifestFullName"></param>      
        /// <returns>bool</returns>
        private static bool Uninstall(string PstrLocalManifestFullName)
        {
            try
            {
               
                string LstrId = GetManifestId(PstrLocalManifestFullName);

                if (LstrId == null) throw new Exception("Invalid manifest file detected.");

                FileInfo LobjFile = new FileInfo(PstrLocalManifestFullName); // grab the local
                Console.WriteLine("Deleting the manifest file: " + PstrLocalManifestFullName);
                LobjFile.Delete();

                Console.WriteLine("Deleting Registry entries...");
                RegistryKey LobjKey = Registry.CurrentUser.OpenSubKey(REG_KEY, true);
                LobjKey.DeleteValue(PstrLocalManifestFullName);
                LobjKey.DeleteSubKey(LstrId);
                Console.WriteLine("Registry keys deleted.");
                Console.WriteLine("Uninstall process completed.");
                return true;
            }
            catch (Exception PobjEx)
            {
                PobjEx.HandleException(true, "Either the manifest path is invalid, cannot be accessed, or " +
                                             "the manifest is not a valid XML file, is not a Web Add-in " +
                                             "manifest, or the local path does not exist, or the local " +
                                             "file is still in use.");
                return false;
            }
        }

        /// <summary>
        /// Updates the manifest from the server location
        /// </summary>
        /// <param name="PstrManifestPath"></param>
        /// <param name="PstrInstallPath"></param>
        /// <returns>bool</returns>
        private static bool Update(string PstrManifestPath, string PstrInstallPath)
        {
            try
            {

                string localManifestFullPath="";

                if (PstrManifestPath.ToLower().StartsWith("http"))
                {
                    localManifestFullPath = DownloadFileFromWeb(PstrManifestPath, PstrInstallPath);
                }
                else
                {
                    localManifestFullPath = DownloadFile(PstrManifestPath, PstrInstallPath);
                }

                Console.WriteLine("Update process completed.");
                return true;
            }
            catch (Exception PobjEx)
            {
                PobjEx.HandleException(true, "Either the manifest server location is invalid, " +
                                             "cannot be accessed, or the local file is in use or " +
                                             "the path is no longer valid.");
                return false;
            }
        }

        /// <summary>
        /// Opens the manifest file, locates the ID and returns it
        /// This is needed for each operation - install / uninstall
        /// </summary>
        /// <param name="PstrFilename"></param>
        /// <returns>string</returns>
        private static string GetManifestId(string LocalManifestPath)
        {
            try
            {
                XmlDocument LobjDoc = new XmlDocument();
                LobjDoc.Load(LocalManifestPath);

                XmlNamespaceManager LobjMgr = new XmlNamespaceManager(LobjDoc.NameTable);
                LobjMgr.AddNamespace(XML_NS, XML_NS_PATH);
                XmlNode LobjIdNode = LobjDoc.DocumentElement.SelectSingleNode(PATH_IdTag, LobjMgr);
                if (LobjIdNode == null) throw new Exception("Id not found in manifest file at " + PATH_IdTag);
                Console.WriteLine("Manifest ID found: " + LobjIdNode.InnerText);
                return LobjIdNode.InnerText;
            }
            catch(Exception PobjEx)
            {
                PobjEx.HandleException(true, "The file provided is not a valid XML file, " +
                                             "or it is not a properly formatted Web Add-in manifest.");
                return null;
            }
        }

        private static string DownloadFile(string srcPath, string destPath)
        {
            FileInfo LobjFile = new FileInfo(srcPath);
            string LstrInstallFilename = Path.Combine(destPath, LobjFile.Name);
               
            Console.WriteLine("Writing the manifest file to the install folder: " + LstrInstallFilename);
            LobjFile.CopyTo(LstrInstallFilename, true);

            return LstrInstallFilename;
        }

        private static string DownloadFileFromWeb(string ServerUrl, string destPath)
        {

            var tmp = ServerUrl.Split('/');
            string manifestName = tmp[tmp.Length - 1];

            string LstrInstallFilename = Path.Combine(destPath, manifestName);

            var request = System.Net.HttpWebRequest.Create(ServerUrl);
            using (var sReader = new StreamReader(request.GetResponse().GetResponseStream()))
            {
                using (var sWriter = new StreamWriter(LstrInstallFilename))
                {
                    sWriter.Write(sReader.ReadToEnd());
                }
            }

            return LstrInstallFilename;
        }
    }
}
