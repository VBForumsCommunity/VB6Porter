1. Download and merge the VBA SDK with WinRAR, for installation.
2. Install the VBA 6.5 SDK
3. Download and unzip both projects
4. Move the VB65 folder into the Template folder, ie C:\Program Files (x86)\Microsoft Visual Studio\VB98\Template
5. Move the files inside the VB98 folder into the MVS folder, ie C:\Program Files (x86)\Microsoft Visual Studio\VB98
6. Compile VBA65.dll to the C:\Program Files (x86)\Microsoft Visual Studio\VB98 folder, as administrator.
7. Compile VB6Porter.dll to the C:\Program Files (x86)\Microsoft Visual Studio\VB98 folder, as administrator.
8. To recompile, you may have to close/release all instances of any loaded IDE/addins and delete/kill the old dll.

This is a comprehensive upgrade for the Visual Basic 6.0 IDE, available as an addin.  The addin was created with the same code base developed for/with the Visual Basic 6.5 IDE/VBA 6.5 SDK integration class.  It duplicates several advanced VB.NET IDE features including, auto indenting (Pretty listing, reformatting of code), scroll wheel support, forms and controls renaming, Start without debugging, greater number of undos, and many code editing functions that prepare vb6 to be more explicit and more interchangable with vb.net standards and the available code base.  A detailed error list can be generated with the Pre-compile button, similar to the "Auto Syntax Check" errors, but with an ability to jump directly to any of the errors listed.  There is also an option to detect file changes outside of the code editor that allows you to reload files from outside, or reject/overwrite them.  It supports a use of the language that accomodates porting VB code, both forwards and backwards. API's and their types can be inserted automatically when the name of the API is typed into the code editor.  Auto class instancing is still in the works to replicate VB.NET namespaces in VB6.

OutLine
* VB6Porter addin dll - yields the VBA object model to developers to combine/port with VB6 object model
* VBA65 Host dll (clsVBA65.cls) - yields the VBA object model to the parent addin above
* VBA65 Host dll (clsAPCVBA65.cls) - yields the  VBA object model to the host class clsVBA65.cls

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

