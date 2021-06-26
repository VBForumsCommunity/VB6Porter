
This is a comprehensive upgrade for the Visual Basic 6.0 IDE, available as an addin.  It supports a use of the language that accomodates porting VB code, both forwards and backwards.  The addin was created with the same code base developed for/with the Visual Basic 6.5 IDE/VBA 6.5 SDK integration class.  It duplicates several advanced VB.NET IDE features including to following options:
* Scroll wheel support
* Auto indenting (Pretty listing)
* Autocompletion of end constructs, ie If/End, Do/Loop, For/Next etc..
* 500 undos/redos
* Start without debugging.  Saves, makes, and starts the built executable
* Modules renaming.  Renames project files to match the module name change
* Many code editing functions that prepare vb6 to be more explicit and more interchangable with vb.net standards and code base
* Detect any project file changes outside of the code editor.  Allows the developer to reload files from outside, or reject/overwrite them
* Project error parsing can be generated with the Pre-compile button, similar to the "Auto Syntax Check" errors.  The developer can jump directly to the listed errors 
* API's and their types can be inserted automatically when the name of the API is typed into the code editor
* Advanced build switches hooked with the intercept method shown by the Trick
* View build output by redirecting the compiler executables and capturing their arguments being passed
* Common VB.NET controls and Namespaces will be available to VB6 later this year, as a user friendly ComVisible NET dll.
* https://github.com/WindowStations/VB6NameSpaces
* https://www.youtube.com/user/primedivine

OutLine
* VB6Porter addin dll - VB6 IDE upgrades with full access/control over the VBA object model, allowing developers to combine/port modules etc between VB6/VBA object models.  VBA 6.5 has some unique features that the extensibility model 5.3/6.0 does not have, including but not limited to, a precompile option and line error parsing abilities.  It serves as a code repository for API delarations, allowing the VB6 code editor instant access to insert code module information into the current development project, from the VBA project space. You may also execute code procedures that are inside of a standard module inside of a coupled VBA project from the VB6 addin.
* VB interop dll - A single VB.NET assembly now makes it possible to port VB.NET code directlly to VB6.  It can also help prepare VB6 source code, for migration to VB.NET.  The assembly interoperates with VB.NET's NameSpaces so that they can consumed by VB6 as nested class buckets.  Significant strides have been made to fully support VB.NET controls/properties/events and Namespaces directly through advanced dynamic interop.
* VBA65 Host dll (clsVBA65.cls) - yields the VBA object model to the parent addin VB6Porter
* VBA65 Host dll (clsAPCVBA65.cls) - yields the  VBA object model to the host class clsVBA65.cls
* (optional) AutoIndenter addin dll - VBA IDE upgrades with the same code editing features and automatic indentation as the VB6Porter addin


Pre-installation
1. Download and merge the VBA SDK with WinRAR, for installation.
2. Install the VBA 6.5 SDK
3. Download and unzip the 3 projects
4. Move the VB65 folder into the Template folder, ie C:\Program Files (x86)\Microsoft Visual Studio\VB98\Template
5. Move the files inside the VB98 folder into the MVS folder, ie C:\Program Files (x86)\Microsoft Visual Studio\VB98
6. Or, rename your C2.EXE to C3.EXE, and LINK.EXE to L1NK.EXE, ie C:\Program Files (x86)\Microsoft Visual Studio\VB98\C3.EXE.  To replace, move OUTARGS\C2.EXE, and OUTARGS\LINK.EXE, into the MVS folder, C:\Program Files (x86)\Microsoft Visual Studio\VB98\C2.EXE
7. Open the OUTARGS\outargs.REG registry file, to redirect Windows to these new files.

Advanced Microsoft Visual Basics (second edition) Chapter 7. Page 275.
Registry edit file (outargs.reg) for C3.EXE, L1NK.EXE output feature:

   >Windows Registry Editor Version 5.00
 
   >[HKEY_CURRENT_USER\Software\VB and VBA Program Settings\LINK\Startup]
   >"RealAppName"="L1NK"

   >[HKEY_CURRENT_USER\Software\VB and VBA Program Settings\C2\Startup]
   >"RealAppName"="C3"


Installation
1. Register the VB.NET assembly to extend VB6 capabilities and to design VB.NET forms with the VB6Porter addin https://github.com/WindowStations/VB6NameSpaces
2.  Compile VBA65.dll to the C:\Program Files (x86)\Microsoft Visual Studio\VB98 folder, as administrator.  This provides access to the VBA object model.
3. Compile VB6Porter.dll to the C:\Program Files (x86)\Microsoft Visual Studio\VB98 folder, as administrator.  This integrates the object models of VBA, VB6 and VB.NET.
4. Compile AutoIndenter.dll to the C:\Program Files (x86)\Microsoft Visual Studio\VB98 folder, as administrator.  This provides the VBA IDE with the same code editing functions from VB6Porter.
* To recompile, you may have to close/release all instances of any loaded IDE/addins and delete/kill the old dll.



Addtional information

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

VBA 6.x includes a multithreading capability (VBAMT) that allows the developer to create multithreaded projects, i.e., projects containing multiple threads that can execute concurrently. These projects are “standalone” in that the code that executes is not tightly coupled to the GUI instance of the application that hosts the VBA IDE. The process begins with a host application that registers one or more multithreading (“MT”) project types. Using these registered project types and an ActiveX VBA MT Designer, users of the application can create multithreaded projects, which are compiled and published as stand-alone DLLs.
A published MT project DLL can be used by other multithreaded host applications. To do so, a thread creates an instance of a VBA MT runtime object and an instance of a global application object associated with the MT project, against which the MT project executes code. The VBA MT runtime object is initialized through the IVbaMT interface, which loads all registered MT project DLLs. The collection of loaded MT project DLLs can be accessed through the IVbaMTDlls interface. Any number of threads can use an MT project DLL concurrently.
When a new MT project is created, a design instance of the MT Designer is created and associated with a logical group of threads. The user can then add classes, forms, and modules to the project. However, VBA MT provides that these classes are not publicly creatable outside of the VBA environment. Nor are these classes instances of the templated objects of an object model. The VBAMT capability provides no mechanisms to assist the end user with integrating a VBAMT project object into an enterprise wide application with a minimum of knowledge by the end user of the specifics of how to integrate into the enterprise environment.
VBA does allow the public creation of an ActiveX designer instance from a standalone project, but ActiveX designers require high-level programming skill and still do not provide the convenient object model integration support that a typical end user would require.

