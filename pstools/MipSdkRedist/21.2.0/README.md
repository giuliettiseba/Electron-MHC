# MipSdkRedist

Used for easy importing of the Milestone Systems MIP SDK components in a PowerShell 5.1 environment.

## Overview

The MIP SDK from Milestone Systems is available as a nuget package. This PowerShell module's single
purpose is to simplify the use of MIP SDK within PowerShell by...

1. Making the content of the MIP SDK available as an easy-to-distribute PowerShell module
2. Loading of the MIP SDK assemblies which, as of 2021 R1, requires a workaround for what you would normally handle using binding redirects. Since it's unreasonable to expect users to add binding redirects to PowerShell's app.config, and attempting to register a handler for `AppDomain.CurrentDomain.AssemblyResolve` within a PowerShell script results in a stack overflow, we're doing this within a `IModuleAssemblyInitializer` in C# instead.

## Installation

### Preferred Method

This module is published on PowerShell Gallery so the recommended installation method is to run the following command...

```powershell
Install-Module MipSdkRedist
```

### Manual Method

If your destination system does not have internet access or you prefer to install the module by hand, you can download the release from PowerShell Gallery or from this repositories Releases page, and extract the ZIP file contents to one of the paths in your `$env:PSModulePath` such as `C:\Program Files\WindowsPowerShell\Modules` or `~\Documents\WindowsPowerShell\Modules`.

Make sure to maintain proper PowerShell module file structure so that the module can be automatically discovered. For example, if you place the module in Program Files where it will be available to all users, the path to the .psd1 file should be `C:\Program Files\WindowsPowerShell\Modules\MipSdkRedist\<version>\MipSdkRedist.psd1` where `<version>` is the module version such as "21.1.3" for the latest version based on MIP SDK 2021 R1.
