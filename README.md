![Logo with represenation of an add-in object sideloaded inside of Office](https://davecra.files.wordpress.com/2020/03/sideload_banner.png)
# WebAddinSideloader
A console application to assist enterprises with Office web add-in side loading in a centralized way.
Office Web Add-ins should be installed using the Office 365 Administration page. However, there are cases where this may not work, or a small team will want to install an add-in for use within that team. This is where Sideloading becomes useful. It is detailed in this article:
[https://docs.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).
This tool assists with sideloading in this method and allows for the ability to not only install, but update and install using this method.

## Download
[Download the latest version of this tool from here.](https://github.com/davecra/WebAddinSideloader/blob/master/Set-WebAddin%20(v1.0.0.1).zip)

## Usage
To use this tool simply run this command:

``` 
Set-WebAddin -help
```

From the help screen you will find the details outlined in the following section.

### Web AddIn Sideloader Command Line Utility
Version: 1.0.0.1
----------------

This utlity is to allow enterprise organizations without Office 365 or centeralized add-in governance to be able to
install web add-ins to users desktops, requiring no effort by the users to have the add-ins installed and available
for use.

Switches
--------
This utility provides the following options:

        -install        Installs the add-in
        -uninstall      Uninstalls the add-in
        -update         Updates the add-in
        -test           Installs the add-in (local only)
        -cleanup        Removes the add-in (local only)

You will also need to provide one or more of these switches with any of the above options:

        -installPath [local folder path]
        -manifestPath [centralized manifest XML file]
        -installedManifestFullname [full path to local manifest] (only with uninstall)

        NOTE: The install path folder MUST exist.

Local Only Testing
------------------
For sideload (local only) testing you can use these switches:

        -test -manifestPath [full path and filename to the manifest*]
        -cleaup -maifestPath [full path and filename to the manifest*]

**NOTE**: The manifest path must be on the local drive.

Once it has been sideloaded, your add-in will not load automatically. To load your add-in you will need to select it from the drop-down list of "Developer Add-ins" in the "My Add-ins" list on the Insert ribbon.

![Developer Add-ins](https://davecra.files.wordpress.com/2020/06/add-in-menu.png?w=400)

Examples
--------
The following are some examples of usage:

 - ``` Set-WebAddin -install -installPath c:\add-in -manifestPath \\server\share\manifest.xml ```
 - ``` Set-WebAddin -install - installPath c:\add-in -manifestPath https://server/path/manifest.xml ```
 - ``` Set-WebAddin -uninstall -installedManifestFullname c:\add-in\manifest.xml ```
 - ``` Set-WebAddin -update -installPath c:\add-in -manifestPath \\server\share\manifest.xml ```
 - ``` Set-WebAddin -test -manifestPath c:\add-in\manifest.xml ```
 - ``` Set-WebAddin -cleanup -manifest c:\add-in\manifest.xml ```
