![script view](https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/demo-full.png "Excel Demo")

# Overview

Here is a simple spreadsheet showing how to access Wasatch Photonics 
spectrometers from inside Microsoft Excel via Wasatch.NET.

The good news is that once you've registered a COM-enabled WasatchNET.DLL
of the correct architecture on your computer, it is extremely easy to
instantiate and manipulate spectrometer objects from Excel macros.  

Click here to see just how simple the Visual Basic code can be:

- https://github.com/WasatchPhotonics/Wasatch.Excel/tree/master/WasatchDemo.vb

The bad news is that a little tweaking may be required to get that DLL built
and registered correctly, but we've tried to hone the process down to a few
easy steps (see [Assembly Notes](#Assembly-Notes), below).

# Dependencies

The Excel demo requires the following DLLs, both available from 
[Wasatch.NET](https://github.com/WasatchPhotonics/Wasatch.NET/tree/master/lib):

* WasatchNET.dll
* LibUsbDotNet.dll

The WasatchNET.dll may need to be custom-compiled to match the architecture of
your Windows OS and your version of Microsoft Office...see 
[Assembly Notes](#Assembly-Notes) below.

It also requires .INF files to associate Wasatch Photonics spectrometers
with LibUSB.NET.  Right now, the easiest way to do that is to install
Enlighten or Dash, one of our standard spectroscopy GUIs.

# Assembly Notes

.NET has changed a bit since Visual Basic 6 (VB6) and Visual Basic for
Applications (VBA) were created. Our Wasatch.NET driver is normally built
for "modern" (.NET 4.0 and above) integrations, using "Any CPU" whenever
possible to avoid architecture (bitness issues).  

Unfortunately, we have to dig back into that nastiness a bit to get things 
working from Excel, but it turns out the process isn't too bad.  We are **not**
leaving Wasatch.NET configured for single-architecture builds by default,
because that would unnecessarily complicate things for more modern 
architecture-neutral platforms; as a result, a custom build of Wasatch.NET may
be required to get things working with your version of Excel on your version of
Windows.  Fortunately, it only takes a few minutes to do (and we'll be glad to
help if you get stuck).

(It's even possible that not all of these steps are required...I can only say 
that this is what I did to get it working for an initial integration.  If you 
find a shorter, simpler or more robust process, please let us know!)

## Build WasatchNET for a specific COM architecture

Normally we build Wasatch.NET for "Any CPU", but it seems to use COM you need
to explicitly build in Visual Studio for either "x86" or "x64", depending on
which architecture of Microsoft Office you're using.

<a href="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/excel-01-architecture.png"><img src="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/excel-01-architecture.png" width="20%" height="20%" align="right"/></a>
Note: this is different from asking whether you're running on a 64-bit version
of Windows or not; many people run 32-bit Office on 64-bit Windows.  To find out
which you're using, run Excel and go to File -> Help.
<br clear="all"/>

<a href="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/visual-studio-01-config-mgr.png"><img src="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/visual-studio-01-config-mgr.png" width="20%" height="20%" align="right"/></a>
You then need to set the same architecture when building Wasatch.NET, using
Build -> Configuration Manager.
<br clear="all"/>

<a href="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/visual-studio-02-com-visible.png"><img src="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/visual-studio-02-com-visible.png" width="20%" height="20%" align="right"/></a>
While you're there, you'll also want to ensure that the WasatchNET library is
"COM Enabled" by right-clicking the WasatchNET project (not Solution), going
to the Application tab, clicking "Assembly" and finally checking "Make assembly
COM-visible".

(You may then have to similarly set the WinFormDemo and Setup installer projects
to x86 or x64 as appropriate.)

### Pre-Built (Win10-64bit, Excel 2010 64-bit)

The version I used for my own testing and the displayed screenshots is provided
in the repository as WasatchNET-x64.zip.  You'll still need to "register" it on your
computer as described below.

## Register the Assembly

After you've obtained or built an appropriate WasatchNET.dll, you'll need to
"register" the assembly on your computer so it can be used by VB6 and VBA.

Recommended process:

- Copy both WasatchNET.DLL and LibUsbDotNet.dll to C:\Windows\System32. 
  Technically they can probably be anywhere in the system %PATH%, so that 
  Excel can find them (DLLs are treated as executables for path purposes).
- Register WasatchNET.DLL via "regasm.exe"
	- open a "cmd" DOS shell using "Run as administrator"
    - run *one* of the following two commands, based on the architecture of 
	  your copy of Microsoft Office (*not* Microsoft Windows)
        - this assumes .NET 4.0 or newer is installed on your computer
        - note the commands differ *only* by the directory "Framework" (x86) vs "Framework64" (x64)

### x86

`C:\> \Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe /codebase /tlb:\Windows\System32\WasatchNET.tlb \Windows\System32\WasatchNET.dll`

### x64

`C:\> \Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe /codebase /tlb:\Windows\System32\WasatchNET.tlb \Windows\System32\WasatchNET.dll`

If it works, you should see something like this:

<pre>
    Microsoft .NET Framework Assembly Registration Utility version 4.7.2046.0
    for Microsoft .NET Framework version 4.7.2046.0
    Copyright (C) Microsoft Corporation.  All rights reserved.

    RegAsm : warning RA0000 : Registering an unsigned assembly with /codebase can cause your assembly to interfere with other applications that may be installed on the same computer. The /codebase switch is intended to be used only with signed assemblies. Please give your assembly a strong name and re-register it.
    Types registered successfully
    Assembly exported to 'C:\Windows\System32\WasatchNET.tlb', and the type library was registered successfully
</pre>

## Enable VBA in Excel 

If you haven't already done so, you need to manually enable the "Developer" ribbon
in Microsoft Excel, which grants access to all the VBA goodies.  (Instructions for
Excel 2010):

<a href="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/excel-02-options.png"><img src="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/excel-02-options.png" width="20%" height="20%" align="right"/></a>
- Start by navigating to File -> Options...
<br clear="all"/>

<a href="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/excel-03-customize-ribbon.png"><img src="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/excel-03-customize-ribbon.png" width="20%" height="20%" align="right"/></a>
- select "Customize Ribbon", then check "Developer"...
<br clear="all"/>

<a href="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/excel-04-toolbar-vb.png"><img src="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/excel-04-toolbar-vb.png" width="20%" height="20%" align="right"/></a>
- adding the "Developer" tab to your ribbon, giving access to the Visual Basic editor...
<br clear="all"/>

<a href="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/excel-05-vba-editor.png"><img src="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/excel-05-vba-editor.png" width="20%" height="20%" align="right"/></a>
- ...providing a functional IDE in the middle of Excel!
<br clear="all"/>

## Add Reference to WasatchNET to your spreadsheet

Now that you have the VBA IDE, we need to add a "reference" to Wasatch.NET so 
that you can refer to its namespace, classes and objects in your code.

<a href="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/ref-01-add.png"><img src="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/ref-01-add.png" width="20%" height="20%" align="right"/></a>
- From the VBA Editor, select Tools->Add Reference...
<br clear="all"/>

<a href="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/ref-02-browse.png"><img src="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/ref-02-browse.png" width="20%" height="20%" align="right"/></a>
- then "Browse" to wherever you copied and registered WasatchNET.dll.  I think 
  you have to actually select the WasatchNET.tlb, which oddly doesn't show its 
  extension in the browse window.
<br clear="all"/>

<a href="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/ref-03-done.png"><img src="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/ref-03-done.png" width="20%" height="20%" align="right"/></a>
- You know you're finished when you see "[x] .NET application wrapper for Wasatch Photonics".

## Get coding!

Using [WasatchDemo.vb](https://github.com/WasatchPhotonics/Wasatch.Excel/tree/master/WasatchDemo.vb) 
as an example, start developing your Office-based spectroscopy application today!

![spreadsheet view](https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/demo-spreadsheet.png "Excel Spreadsheet")
