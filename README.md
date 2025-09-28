Updates coming this [December 2025!](https://youtu.be/cfCNbrJlw84?si=TmaGRzmPje1JWy2l)
* Much tigher integration between the form designer and vb6 code modules.
* Cleaned up some stray underscores and used [] to avoid collisions between keywords and properties or method names.
* Updated the vb.net module to yield the default event when spawning them from the vb6 IDE.
* Completed the full set of vb.net controls and components.
* Updated the save/load functions to automatically save/load.
* Code formatting enhancments.
* Backspace over white spaces while adhering to the code format.
* Proper vb6 code folding, the right way.
* In line autocompletion optional.  Many syntax errors starting with 'Expected X': can be autocompleted with spacebar.  It saves time typing to construct the same line.
* For instance, Type Dim i, then press spacebar.  It will generate Dim i As , with the intellisense open to select the type with arrow key.
* The 'Open' file constructor is optional.  Type Open, then press spacebar.  The whole file open and close structure will appear with randomized file number.
* The 'msgbox' result constructor is optional.  Type: msgbox( Then press: spacebar.  The whole msgbox constructor will appear.  It generates a randomized variable for the yes/no response.
* On Error completion is helpful. The Goto option constructs the whole error handling snippet including a jump line and Erl with description.  It avoids the exit sub/function keywords.  It allows additional error handling of the procedure, ie On Error Resume Next can still be used before or after the GoTo/Exception region.  Simply type: On Error spacebar.
* Line numbers are added dynamically at compile time, and erased by design time.  The On Error constructor now has the line number from ERL.
* Fixed bug where drag/dropping code caused extra spaces to be generated when using line numbers.
* Added kill button with option to kill all processes (if more than one) matching the same name as the started process.
* Added the kill function to the Start button process, ensuring startup is handled properly if tested application is hanging in the background.
* Export detailed project info in XML format similar to VB.NET.  Save file to desired location.
* Insert XML header to a method similar to VB.NET.  Type ''' above any method or property.
* Sound queues optional.
* More to come...

![alt text](https://user-images.githubusercontent.com/39764372/137580245-97a931ea-6382-4400-9822-b850e0eb6603.png)
![venn](https://github.com/user-attachments/assets/cfbe7652-9edb-455d-8c6c-65070af327c1)

https://www.youtube.com/watch?v=RJmDeGAroR4
This is a comprehensive upgrade for the Visual Basic 6.0 IDE, available as an addin.  It supports a use of the language that accommodates porting VB code, both forwards and backwards.  The addin was created with the same code base developed for/with the Visual Basic 6.5 IDE/VBA 6.5 SDK integration class.  It duplicates several advanced VB.NET IDE features including the following options:
* Scroll wheel support
* Auto indenting (Pretty listing)
* Autocompletion of end constructs, ie If/End, Do/Loop, For/Next etc..
* 500 undos/redos
* Detect any project file changes outside of the code editor.  Allows the developer to reload files from outside, or reject/overwrite them
* Start without debugging.  Saves, makes, and starts the built executable
* Modules renaming.  Renames project files to match the module name change
* Code editing functions that prepare vb6 to be more explicit and more interchangable with vb.net standards and code base
* Advanced build switches hooked with the intercept method shown by the Trick
* View build output by redirecting the compiler (vb6.exe to c2.exe & link.exe) and capturing the arguments being passed between them
* VB.NET Forms, common controls, and Namespaces with a modern form/control designer
* Visual Basic 6.5 (2007) icon pack

OutLine
* VB interop dll - A single VB.NET assembly now makes it possible to backport VB.NET code directly to VBA/VB6, or it can help migrate VB6 code to VB.NET.  The assembly interoperates with VB.NET's NameSpaces so that they can be consumed by VB6 as nested class buckets.  Significant strides have been made to fully support VB.NET controls/properties/events and Namespaces directly through advanced dynamic interop.
* VB6Porter addin dll - VB6 IDE upgrades with full access and control over the object models of VBA and VB.NET, along with the interop assembly (VB.DLL).  It allows developers to combine/port modules etc between VB6/VBA/VB.NET. 
* (Optional) - VBAPorter addin dll - VBA IDE upgrades similar to VB6Porter above.
* (Optional) - VBA65 Host dll (clsVBA65.cls) yields the VBA object model to the parent addin VB6Porter - VBA65 Host dll (clsAPCVBA65.cls) yields the  VBA object model to the host class clsVBA65.cls.  VBA 6.5 has some unique features that the extensibility model 5.3/6.0 does not have, including but not limited to, a precompile option and line error parsing abilities.  It serves as a code repository for API declarations, allowing the VB6 code editor instant access to insert code module information into the current development project, from the VBA project space. You may also execute code procedures that are inside of a standard module inside of a coupled VBA project from the VB6 addin. Nebcore software solutions LLC licensed VBA 6.4-6.5. https://www.nhcompanyregistry.com/companies/nebcore/


Pre-installation and options
1. Download and unzip the project files.
2. Rename your C2.EXE to C3.EXE, and LINK.EXE to L1NK.EXE, ie C:\Program Files (x86)\Microsoft Visual Studio\VB98\C3.EXE.
3. Replace the files with the project files from OUTARGS, ie OUTARGS\C2.EXE, and OUTARGS\LINK.EXE
4. Open the OUTARGS\outargs.REG registry file, to redirect Windows to these new files.

Advanced Microsoft Visual Basics (second edition) Chapter 7. Page 275.
Registry edit file (outargs.reg) for C3.EXE, L1NK.EXE output feature:
```
   Windows Registry Editor Version 5.00
 
   [HKEY_CURRENT_USER\Software\VB and VBA Program Settings\LINK\Startup]
   "RealAppName"="L1NK"

   [HKEY_CURRENT_USER\Software\VB and VBA Program Settings\C2\Startup]
   "RealAppName"="C3"
```


Installation 

* Installer: https://github.com/VBForumsCommunity/VB6Porter/raw/master/SetupVB6Porter.exe


method 2
* Register the VB.NET assembly named VB.dll with the setup file named SetupRegisterAssembly.exe https://github.com/WindowStations/VB6NameSpaces
* Compile or register VB6Porter.dll as administrator, inside the folder location: C:\Program Files (x86)\Microsoft Visual Studio\VB98.  This integrates the object models of VBA, VB6 and VB.NET.  * To re-compile, you may have to close/release all instances of any loaded IDE/addins and delete/kill the old dll.



