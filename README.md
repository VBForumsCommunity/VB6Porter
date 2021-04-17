
OutLine
* VB6Porter addin dll - VB6 IDE upgrades with full access/control over the VBA object model, allowing developers to combine/port modules etc between VB6/VBA object models.  VBA 6.5 has some unique features that the extensibility model 5.3/6.0 does not have, including but not limited to, a precompile option and line error parsing abilities.
* VBA65 Host dll (clsVBA65.cls) - yields the VBA object model to the parent addin VB6Porter
* VBA65 Host dll (clsAPCVBA65.cls) - yields the  VBA object model to the host class clsVBA65.cls

This is a comprehensive upgrade for the Visual Basic 6.0 IDE, available as an addin.  The addin was created with the same code base developed for/with the Visual Basic 6.5 IDE/VBA 6.5 SDK integration class.  It duplicates several advanced VB.NET IDE features including, auto indenting (Pretty listing, reformatting of code), scroll wheel support, forms and controls renaming, Start without debugging, greater number of undos, and many code editing functions that prepare vb6 to be more explicit and more interchangable with vb.net standards and the available code base.  A detailed error list can be generated with the Pre-compile button, similar to the "Auto Syntax Check" errors, but with an ability to jump directly to any of the errors listed.  There is also an option to detect file changes outside of the code editor that allows you to reload files from outside, or reject/overwrite them.  It supports a use of the language that accomodates porting VB code, both forwards and backwards. API's and their types can be inserted automatically when the name of the API is typed into the code editor.  Auto class instancing is still in the works to replicate VB.NET namespaces in VB6.

1. Download and merge the VBA SDK with WinRAR, for installation.
2. Install the VBA 6.5 SDK
3. Download and unzip the 3 projects
4. Move the VB65 folder into the Template folder, ie C:\Program Files (x86)\Microsoft Visual Studio\VB98\Template
5. Move the files inside the VB98 folder into the MVS folder, ie C:\Program Files (x86)\Microsoft Visual Studio\VB98
5b. Or, rename your C2.EXE to C3.EXE, and LINK.EXE to L1NK.EXE, ie C:\Program Files (x86)\Microsoft Visual Studio\VB98\C3.EXE
6. To replace, move OUTARGS\C2.EXE, and OUTARGS\LINK.EXE, into the MVS folder, C:\Program Files (x86)\Microsoft Visual Studio\VB98\C2.EXE
7. Open the OUTARGS\outargs.REG registry file, to redirect Windows to the new files.
8. Compile VBA65.dll to the C:\Program Files (x86)\Microsoft Visual Studio\VB98 folder, as administrator.
9. Compile VB6Porter.dll to the C:\Program Files (x86)\Microsoft Visual Studio\VB98 folder, as administrator.
10. To recompile, you may have to close/release all instances of any loaded IDE/addins and delete/kill the old dll.

Advanced Microsoft Visual Basics (second edition) Chapter 7. Page 275.
Registry edit file (outargs.reg) for C3.EXE, L1NK.EXE output feature:

>Windows Registry Editor Version 5.00
 
>[HKEY_CURRENT_USER\Software\VB and VBA Program Settings\LINK\Startup]
>"RealAppName"="L1NK"

>[HKEY_CURRENT_USER\Software\VB and VBA Program Settings\C2\Startup]
>"RealAppName"="C3"


 Visual Basic for Applications is an embeddable BASIC language software product comprised of the following components:
 * VBA
 * APC
 * Microsoft Forms
 * Core Technology   
 * End User Documentation and derivatives
 * VBA Core Installer Package

VBA 
These are the binary files that make up the core VBA deliverable.  This includes the language runtime and user interface components including editing, debugging, project management, property control. It also includes the files need for Multi-threading and the multi-threading runtime. 
VBE6.DLL, VBE6EXT.OLB, SCP32.DLL, VBE6INTL.DLL, VBAME.DLL, LINK.EXE, MSPDB60.DLL, MTDSR.DLL, VBA6MTRT.DLL, VB6DEBUG.DLL 
  
APC 
These are the binary files that make up the application programming interface (API) layer to integrate VBA.
APC65.DLL, APC60ITL.DLL 

Microsoft Forms
This provides a complete visual editing and dialog design environment. 
FM20.DLL, FM20ENU.DLL, RICHED20.DLL

Core Technology
VBA is dependent on a set of MICROSOFT core proprietary technology.  These technologies are to be consumed by VBA only.  Hosts of VBA are forbidden from accessing core technology directly. 
MSO.DLL, MSOINTL.DLL, SELFCERT.EXE, SIGNER.DLL, MSVBVM60.DLL, MSSTDFMT.DLL, MSSTKPRP.DLL, HLP95EN.DLL    

VBA Core Installer Package (MSI)
MICROSOFT provides a setup package that installs all the core VBA technologies onto an end users machine.  This prevents accidental disabling of VBA functionality due to an improper installation of VBA components. 

