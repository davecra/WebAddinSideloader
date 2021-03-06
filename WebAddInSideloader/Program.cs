﻿using Microsoft.Win32;
using System;
using System.IO;
using System.Net;
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
                if (LobjArgs.Test) LbolSuccess = Test(LobjArgs.ManifestPath);
                if (LobjArgs.Cleanup) LbolSuccess = Uninstall(LobjArgs.ManifestPath, true);
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
                   "Switches\n" +
                   "--------\n" +
                   "This utility provides the following options: \n\n" +
                   "\t-install\tInstalls the add-in\n" +
                   "\t-uninstall\tUninstalls the add-in\n" +
                   "\t-update\t\tUpdates the add-in\n" +
                   "\t-test\t\tInstalls the add-in (local only)\n" +
                   "\t-cleanup\t\tRemoves the add-in (local only)\n\n" +
                   "You will also need to provide one or more of these switches with any of the above options: \n\n" +
                   "\t-installPath [local folder path]\n" +
                   "\t-manifestPath [centralized manifest XML file]\n" +
                   "\t-installedManifestFullname [full path to local manifest] (only with uninstall)\n\n" +
                   "\tNOTE: The install path folder MUST exist. \n\n" +
                   "Local Only Testing\n" +
                   "------------------\n" +
                   "For sideload (local only) testing you can use these switches:\n\n" +
                   "\t-test -manifestPath [full path and filename to the manifest*]\n" +
                   "\t-cleaup -manifestPath [full path and filename to the manifest*]\n\n" +
                   "*NOTE: The manifest path must be on the local drive.\n\n"+
                   "Examples\n" +
                   "--------\n" +
                   "The following are some examples of usage: \n\n" +
                   "Set-WebAddin -install -installPath c:\\add-in -manifestPath \\\\server\\share\\manifest.xml\n" +
                   "Set-WebAddin -install - installPath c:\\add-in -manifestPath https://server/path/manifest.xml \n" +
                   "Set-WebAddin -uninstall -installedManifestFullname c:\\\\add-in\\manifest.xml" +
                   "Set-WebAddin -update -installPath c:\\add-in -manifestPath \\\\server\\share\\manifest.xml\n" +
                   "Set-WebAddin -test -manifestPath c:\\add-in\\manifest.xml\n" +
                   "Set-WebAddin -cleanup -manifestPath c:\\add-in\\manifest.xml";
        }

        /// <summary>
        /// Sideloads the add-in for testing from the location of the manifest
        /// Does not download and install, just uses the manifest in-place
        /// </summary>
        /// <param name="PstrManifestPath"></param>
        /// <returns></returns>
        private static bool Test(string PstrManifestPath)
        {
            try
            {
                string LstrId = GetManifestId(PstrManifestPath);
                if (LstrId == null) throw new Exception("Invalid manifest file detected.");
                Console.WriteLine("Writing Registry entries...");
                RegistryKey LobjKey = Registry.CurrentUser.OpenSubKey(REG_KEY, true);
                if (LobjKey == null) LobjKey = Registry.CurrentUser.CreateSubKey(REG_KEY, true);
                if (LobjKey == null) throw new Exception("Unable to create or write to the registry path: HKCU\\" + REG_KEY);
                LobjKey.SetValue(LstrId, PstrManifestPath, RegistryValueKind.String);
                LobjKey = LobjKey.CreateSubKey(LstrId);
                LobjKey.SetValue(VALUE_UseDirectDebugger, 1, RegistryValueKind.DWord);
                LobjKey.SetValue(VALUE_UseLiveReload, 0, RegistryValueKind.DWord);
                LobjKey.SetValue(VALUE_UseWebDebugger, 0, RegistryValueKind.DWord);
                Console.WriteLine("Registry keys completed.");
                Console.WriteLine("Test sideload process completed.");
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

                string LstrLocalManifestFullPath = "";

                if (PstrManifestPath.ToLower().StartsWith("http"))
                {
                    LstrLocalManifestFullPath = DownloadFileFromWeb(PstrManifestPath, PstrInstallPath);
                }
                else
                {
                    LstrLocalManifestFullPath = DownloadFile(PstrManifestPath, PstrInstallPath);
                }

                Console.WriteLine("Copy complete. Analyzing file...");
                string LstrId = GetManifestId(LstrLocalManifestFullPath);
                
                if (LstrId == null) throw new Exception("Invalid manifest file detected.");
                Console.WriteLine("Writing Registry entries...");
                RegistryKey LobjKey = Registry.CurrentUser.OpenSubKey(REG_KEY, true);
                if (LobjKey == null) LobjKey = Registry.CurrentUser.CreateSubKey(REG_KEY, true);
                if (LobjKey == null) throw new Exception("Unable to create or write to the registry path: HKCU\\" + REG_KEY);
                LobjKey.SetValue(LstrId, LstrLocalManifestFullPath, RegistryValueKind.String);
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
        private static bool Uninstall(string PstrLocalManifestFullName, bool PbolDoNotDelete = false)
        {
            try
            {
                string LstrId = GetManifestId(PstrLocalManifestFullName);

                if (LstrId == null) throw new Exception("Invalid manifest file detected.");

                if (!PbolDoNotDelete)
                {
                    FileInfo LobjFile = new FileInfo(PstrLocalManifestFullName); // grab the local
                    Console.WriteLine("Deleting the manifest file: " + PstrLocalManifestFullName);
                    LobjFile.Delete();
                }

                Console.WriteLine("Deleting Registry entries...");
                RegistryKey LobjKey = Registry.CurrentUser.OpenSubKey(REG_KEY, true);
                LobjKey.DeleteValue(LstrId);
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

                string LstrLocalManifestFullPath="";

                if (PstrManifestPath.ToLower().StartsWith("http"))
                {
                    LstrLocalManifestFullPath = DownloadFileFromWeb(PstrManifestPath, PstrInstallPath);
                }
                else
                {
                    LstrLocalManifestFullPath = DownloadFile(PstrManifestPath, PstrInstallPath);
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

        /// <summary>
        /// Downloads the file from the local file system or from a 
        /// network location via UNC
        /// </summary>
        /// <param name="PstrSrcPath"></param>
        /// <param name="PstrDestPath"></param>
        /// <returns></returns>
        private static string DownloadFile(string PstrSrcPath, string PstrDestPath)
        {
            try
            {
                FileInfo LobjFile = new FileInfo(PstrSrcPath);
                string LstrInstallFilename = Path.Combine(PstrDestPath, LobjFile.Name);
                Console.WriteLine("Reading manifest file from: " + PstrSrcPath);
                Console.WriteLine("Writing the manifest file to the install folder: " + LstrInstallFilename);
                LobjFile.CopyTo(LstrInstallFilename, true);

                return LstrInstallFilename;
            }
            catch(Exception PobjEx)
            {
                throw PobjEx.PassException("Unable to download the file.");
            }
        }

        /// <summary>
        /// Downloads the manifest file from the web
        /// </summary>
        /// <param name="PstrServerUrl"></param>
        /// <param name="PstrDestPath"></param>
        /// <returns></returns>
        private static string DownloadFileFromWeb(string PstrServerUrl, string PstrDestPath)
        {
            try
            {
                string[] LstrPathParts = PstrServerUrl.Split('/');
                string LstrManifestName = LstrPathParts[LstrPathParts.Length - 1];

                string LstrInstallFilename = Path.Combine(PstrDestPath, LstrManifestName);
                Console.WriteLine("Reading manifest file from: " + PstrServerUrl);
                Console.WriteLine("Writing the manifest file to the install folder: " + LstrInstallFilename);
                WebRequest LobjRequest = System.Net.HttpWebRequest.Create(PstrServerUrl);
                using (StreamReader LobjReader = new StreamReader(LobjRequest.GetResponse().GetResponseStream()))
                {
                    using (StreamWriter LobjWriter = new StreamWriter(LstrInstallFilename))
                    {
                        LobjWriter.Write(LobjReader.ReadToEnd());
                    }
                }

                return LstrInstallFilename;
            }
            catch(Exception PobjEx)
            {
                throw PobjEx.PassException("Unable to download manifest from web.");
            }
        }
    }
}
