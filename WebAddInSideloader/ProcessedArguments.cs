using System;
using System.IO;

namespace WebAddInSideloader
{
    public class ProcessedArguments
    {
        /// <summary>
        /// ctor - INIT class object with values
        /// </summary>
        /// <param name="PobjArgs"></param>
        public ProcessedArguments(string[] PobjArgs)
        {
            try
            {
                int LintArgCount = 0;
                foreach (string LstrRawArg in PobjArgs)
                {
                    // strip off all leading dashes or slashes
                    string LstrArg = LstrRawArg.StartsWith("-") || LstrRawArg.StartsWith("/") ? 
                                     LstrRawArg.Substring(1) : 
                                     LstrRawArg;
                    // now process
                    if (LstrArg.ToLower() == "?" ||
                        LstrArg.ToLower() == "help")
                    {
                        Help = true;
                        return;
                    }
                    if (LstrArg.ToLower() == "test") Test = true;
                    if (LstrArg.ToLower() == "cleanup") Cleanup = true;
                    if (LstrArg.ToLower() == "install") Install = true;
                    if (LstrArg.ToLower() == "update") Update = true;
                    if (LstrArg.ToLower() == "uninstall") Uninstall = true;
                    if (LstrArg.ToLower() == "manifestpath")
                    {
                        // verify that the path to the manifest is valid and is an XML file

                        var tmp = PobjArgs[LintArgCount + 1].ToLower();

                        if (tmp.StartsWith("http") || ConfirmPath(tmp, true))
                        {
                            if (tmp.EndsWith("xml"))
                            {
                                ManifestPath = tmp;
                            }
                            else
                            {
                                throw new Exception("Manifest path must be an xml file.");
                            }                               
                        }
                        else
                        {
                            throw new Exception("An invalid ManifestPath was specified.");
                        }
                    }
                    // verify that the path to install is a valid folder location
                    if(LstrArg.ToLower() == "installpath")
                    {
                        if(ConfirmPath(PobjArgs[LintArgCount + 1]))
                        {
                            InstallPath = PobjArgs[LintArgCount + 1];
                        }
                        else
                        {
                            throw new Exception("An invalid InstallPath was specified.");
                        }
                    }

                    if (LstrArg.ToLower() == "installedmanifestfullname")
                    {
                        if (ConfirmPath(PobjArgs[LintArgCount + 1], true))
                        {
                            InstalledManifestFullName = PobjArgs[LintArgCount + 1];
                        }
                        else
                        {
                            throw new Exception("An invalid Installed Manifest was specified.");
                        }
                    }
                    LintArgCount++;
                }
                
                // validate the install/update/uninstall/test/cleanup switches
                if ((Install == true && Uninstall == true) ||
                   (Install == true && Update == true) ||
                   (Update == true && Uninstall == true) ||
                   (Install == true && Test == true) ||
                   (Uninstall == true && Test == true) ||
                   (Update == true && Test == true) ||
                   (Cleanup == true && Test == true) ||
                   (Cleanup == true && Uninstall == true) ||
                   (Cleanup == true && Update == true) ||
                   (Cleanup == true && Install == true))
                    throw new Exception("Invalid switch combination.");

                if (Uninstall == true)
                {
                   if (string.IsNullOrEmpty(InstalledManifestFullName))
                    {
                        throw new Exception("Installed Manifiest location must be secpified for uninstall.");
                    }
                       
                }
                else if(Install == true || Update == true)  //this is an install or update
                {
                    // validate that paths are specified
                    if (string.IsNullOrEmpty(ManifestPath) || string.IsNullOrEmpty(InstallPath))
                    {
                        throw new Exception("Both a ManifestPath and an InstallPath must be specified.");
                    }
                }
                else if(Test == true || Cleanup == true) // this is for testing/local only
                {
                    if(string.IsNullOrEmpty(ManifestPath))
                    {
                        throw new Exception("ManifestPath must be specified.");
                    }
                    if(!string.IsNullOrEmpty(InstallPath))
                    {
                        throw new Exception("The install path is not required for this option.");
                    }
                }
                 
                // all is good
                Processed = true;
            }
            catch(Exception PobjEx)
            {
                PobjEx.HandleException(true, "The arguments were invalid. Please use -help for guidance.");
            }
        }

        /// <summary>
        /// Confirms the path specified
        /// </summary>
        /// <param name="PstrPath"></param>
        /// <param name="PbolIsFile"></param>
        /// <returns></returns>
        private bool ConfirmPath(string PstrPath, bool PbolIsFile = false)
        {
            try
            {
                if (PbolIsFile) return new FileInfo(PstrPath).Exists;
                else return new DirectoryInfo(PstrPath).Exists;
            }
            catch(Exception PobjEx)
            {
                PobjEx.HandleException(true);
                return false;
            }
        }

        /// <summary>
        /// PROPERTIES
        /// </summary>
        public bool Help { get; private set; }
        public bool Processed { get; private set; } = false;
        public bool Install { get; private set; }
        public string ManifestPath { get; private set; }
        public string InstallPath { get; private set; }
        public bool Update { get; private set; }
        public bool Uninstall { get; private set; }
        public string InstalledManifestFullName { get; set; }
        public bool Test { get; private set; }
        public bool Cleanup { get; private set; }
    }
}
