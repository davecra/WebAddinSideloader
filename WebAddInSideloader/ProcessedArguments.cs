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
                    if (LstrArg.ToLower() == "install") Install = true;
                    if (LstrArg.ToLower() == "update") Update = true;
                    if (LstrArg.ToLower() == "uninstall") Uninstall = true;
                    if (LstrArg.ToLower() == "manifestpath")
                    {
                        // verify that the path to the manifest is valid and is an XML file
                        if (ConfirmPath(PobjArgs[LintArgCount + 1], true) && 
                            PobjArgs[LintArgCount + 1].ToLower().EndsWith("xml"))
                        {
                            ManifestPath = PobjArgs[LintArgCount + 1];
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
                    LintArgCount++;
                }
                // validate the install/update/uninstall switches
                if ((Install == true && Uninstall == true) ||
                    (Install == true && Update == true) || 
                    (Update==true && Uninstall == true)) 
                    throw new Exception("Invalid switch combination.");
                // validate that paths are specified
                if (string.IsNullOrEmpty(ManifestPath) || string.IsNullOrEmpty(InstallPath))
                    throw new Exception("Both a ManifestPath and an InstallPath must be specified.");
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
    }
}
