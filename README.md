![VBA Editor](https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/demo-full.png)

# Overview

Here is a simple spreadsheet showing how to access Wasatch Photonics 
spectrometers from inside Microsoft Excel via Wasatch.NET.

Click here to see just how simple the Visual Basic code can be:

- https://github.com/WasatchPhotonics/Wasatch.Excel/tree/master/WasatchDemo.vb

# Dependencies

The Excel demo requires that [Wasatch.NET](https://github.com/WasatchPhotonics/Wasatch.NET/)
be installed, so if you haven't done so, download and run the appropriate installer here:

* http://wasatchphotonics.com/binaries/drivers/Wasatch.NET/

A few notes to ensure faultless installation:

<a href="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/excel-01-architecture.png"><img src="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/excel-01-architecture.png" width="20%" height="20%" align="right"/></a>
1. Note that You'll need to download the Wasatch.NET architecture (bitness) 
   corresponding to *your copy of Excel*, which may be different than your 
   version of Windows.  (Many companies install 32-bit Office onto 64-bit Windows
   by default, and unless you're doing VBA programming you may never have reason
   to know the difference.) Use File -> Help to check which version of Excel 
   you're using.
<br clear="all"/>

2. Remember to configure the libusb drivers the first time you connect a Wasatch
   Photonics spectrometer to a computer:

- https://github.com/WasatchPhotonics/Wasatch.NET#post-install-step-1-libusb-drivers

3. The Wasatch.NET installers strive to do everything for you...but one thing 
   they don't currently do is register the COM assembly so it can be used by VB6
   and VBA.  For that, you need to lastly run the RegisterCOM.bat script as 
   administrator, as described here:

- https://github.com/WasatchPhotonics/Wasatch.NET#post-install-step-2-com-registration-optional

That's it! Not such a heavy price to pay in order to perform live spectroscopy 
from a spreadsheet, is it :-)

# Support

If you have any issues, please let us know and we'll do our best to resolve them
expediently:

    support@wasatchphotonics.com

# Appendix: Excel VBA Quick Start

New to Visual Basic for Applications (VBA)?  Here are some quick steps to get you going!

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
- From the VBA Editor, select Tools->References...
<br clear="all"/>

<a href="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/ref-02-browse.png"><img src="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/ref-02-browse.png" width="20%" height="20%" align="right"/></a>
- then "Browse" to C:\\Windows\\WasatchNET.tlb (note: *not* the .DLL next to it!)
<br clear="all"/>

<a href="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/ref-03-done.png"><img src="https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/ref-03-done.png" width="20%" height="20%" align="right"/></a>
- You know you're finished when you see "[x] .NET application wrapper for Wasatch Photonics".

## Get coding!

Using [WasatchDemo.vb](https://github.com/WasatchPhotonics/Wasatch.Excel/tree/master/WasatchDemo.vb) 
as an example, and refering to our [API documentation](http://www.wasatchphotonics.com/api/Wasatch.NET/) 
as needed, start developing your Office-based spectroscopy application today!

![Spreadsheet](https://github.com/WasatchPhotonics/Wasatch.Excel/raw/master/screenshots/demo-spreadsheet.png)
