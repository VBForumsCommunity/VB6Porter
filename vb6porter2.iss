#define MyAppName "VB6Porter"
#define MyAppVersion "2.0"
#define MyAppPublisher "Window Stations"
#define MyAppURL "https://github.com/VBForumsCommunity/VB6Porter"
#define MyAppExeName1 "VB6.EXE"
#define MyAppExeName2 "C2.EXE"
#define MyAppExeName3 "LINK.EXE"

[Setup]
AppId={{63870773-C7EF-4A94-A8D6-F80AB4812363}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\Microsoft Visual Studio
DisableDirPage=yes
DisableProgramGroupPage=yes
LicenseFile=EULA.TXT
;PrivilegesRequired=lowest
OutputBaseFilename=SetupVB6Porter
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ChangesAssociations=yes
SetupIconFile=vbporter64.ico
DiskSpanning=yes
DiskSliceSize=15000000
DisableFinishedPage=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"


[Tasks]
Name: licenseVB6Porter; Description: VB6Porter & VB6Namespaces (VBANET.dll); GroupDescription: "Add-in extensibility"
Name: licenseFull; Description: "Use the newest updated files for Microsoft Visual Basic 2007"; GroupDescription: "License type:"; Flags: unchecked
Name: licenseFull\licenseLearning; Description: "Learning edition"; GroupDescription: "License type:"; Flags: exclusive unchecked
Name: licenseFull\licensePro; Description: "Professional edition"; GroupDescription: "License type:"; Flags: exclusive unchecked
Name: licenseFull\licenseEnt; Description: "Enterprise edition"; GroupDescription: "License type:"; Flags: exclusive unchecked


[Files]
;System   onlyifdestfileexists
Source: "system\lic.reg"; DestDir: {sys}; Flags: restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull; AfterInstall: InstallProgramB
Source: "msstdfmt.dll"; DestDir: {win}\SysWOW64; Flags: restartreplace uninsneveruninstall overwritereadonly regserver; Tasks: licensefull
Source: "tlbinf32.dll"; DestDir: {win}\System32; Flags: restartreplace uninsneveruninstall overwritereadonly regserver; Tasks: licensefull
Source: "OLEPRO32.dll"; DestDir: {sys}; Flags: restartreplace uninsneveruninstall overwritereadonly regserver; Tasks: licensefull


;directory structure
Source: "system\*"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall overwritereadonly recursesubdirs; Tasks: licensefull
Source: "systemregserver\*"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall overwritereadonly recursesubdirs regserver; Tasks: licensefull 
Source: "systemregtlib\*"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall overwritereadonly recursesubdirs regtypelib; Tasks: licensefull 
Source: "shared\*"; DestDir: "{commonpf32}\Common Files\Microsoft Shared"; Flags: restartreplace uninsneveruninstall overwritereadonly recursesubdirs; Tasks: licensefull
Source: "sharedregister\*"; DestDir: "{commonpf32}\Common Files\Microsoft Shared"; Flags: restartreplace uninsneveruninstall overwritereadonly recursesubdirs regserver noregerror; Tasks: licensefull
Source: "sharedregtlib\*"; DestDir: "{commonpf32}\Common Files\Microsoft Shared"; Flags: restartreplace uninsneveruninstall overwritereadonly recursesubdirs regtypelib noregerror; Tasks: licensefull
Source: "common\HOSTWIZ.DLL"; DestDir: "{app}\Common"; Flags: restartreplace uninsneveruninstall overwritereadonly recursesubdirs regserver; Tasks: licensefull 
Source: "common\SUBWIZ.TLB"; DestDir: "{app}\Common"; Flags: restartreplace uninsneveruninstall overwritereadonly recursesubdirs regtypelib; Tasks: licensefull 

;Program files(x86)
Source: "MSADDNDR.dep"; DestDir: {commonpf32}\Common Files\Designer; Flags: restartreplace uninsneveruninstall overwritereadonly
Source: "MSADDNDR.dll"; DestDir: {commonpf32}\Common Files\Designer; Flags: restartreplace uninsneveruninstall overwritereadonly regserver
Source: "MSADDNDR.olb"; DestDir: {commonpf32}\Common Files\Designer; Flags: restartreplace uninsneveruninstall overwritereadonly
Source: "MSADDNDR.tlb"; DestDir: {commonpf32}\Common Files\Designer; Flags: restartreplace uninsneveruninstall overwritereadonly regtypelib
Source: "DAO360.dll"; DestDir: {commonpf32}\Common Files\Microsoft Shared\DAO; Flags: restartreplace uninsneveruninstall overwritereadonly regserver; Tasks: licensefull
Source: "LINK.exe"; DestDir: {commonpf32}\Common Files\Microsoft Shared\VBA\VBA6; Flags: restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull 
Source: "MSPDB60.dll"; DestDir: {commonpf32}\Common Files\Microsoft Shared\VBA\VBA6; Flags: ignoreversion restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull 
Source: "VB6DEBUG.dll"; DestDir: {commonpf32}\Common Files\Microsoft Shared\VBA\VBA6; Flags: restartreplace uninsneveruninstall overwritereadonly regserver; Tasks: licensefull
Source: "vba6mtrt.dll"; DestDir: {commonpf32}\Common Files\Microsoft Shared\VBA\VBA6; Flags: restartreplace uninsneveruninstall overwritereadonly regserver; Tasks: licensefull
Source: "VBE6.dll"; DestDir: {commonpf32}\Common Files\Microsoft Shared\VBA\VBA6; Flags: restartreplace uninsneveruninstall overwritereadonly regserver; Tasks: licensefull
Source: "VBE6EXT.olb"; DestDir: {commonpf32}\Common Files\Microsoft Shared\VBA\VBA6; Flags: restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull
Source: "dataview.dll"; DestDir: {commonpf32}\Common Files\Microsoft Shared\VBA\VBA6\; Flags: restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull  
source: "VBE6INTL.dll"; DestDir: {commonpf32}\Common Files\Microsoft Shared\VBA\VBA6\1033; Flags: restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull 

Source: "C2.exe"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull 
Source: "CVPACK.exe"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull 
Source: "DAO350.dll"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace uninsneveruninstall overwritereadonly regserver; Tasks: licensefull
Source: "DAO2535.tlb"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull 
Source: "LINK.exe"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull
Source: "MRT7ENU.dll"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull
Source: "MSDIS110.dll"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull 
Source: "MSO97RT.dll"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull
Source: "MSPDB60.dll"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull 

Source: "VB6.exe"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull
Source: "VB6Porter.exe"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull

Source: "VB6.olb"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull
Source: "VB6DEBUG.dll"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace uninsneveruninstall overwritereadonly regserver; Tasks: licensefull
Source: "VB6EXT.olb"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull 
Source: "VB6IDE.dll"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull
Source: "VBA6.dll"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull
Source: "VBAEXE6.lib"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull 

Source: "Interop.MSAPC.dll"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull
Source: "Interop.Office.dll"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull
Source: "Interop.VBA65.dll"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull
Source: "Interop.VBIDE.dll"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace uninsneveruninstall overwritereadonly; Tasks: licensefull

Source: "Template\*"; DestDir: "{app}\VB98\Template"; Flags: restartreplace uninsneveruninstall overwritereadonly recursesubdirs; Tasks: licensefull 
Source: "Wizards\*"; DestDir: "{app}\VB98\Wizards"; Flags: restartreplace uninsneveruninstall overwritereadonly recursesubdirs; Tasks: licensefull 
Source: "Wizardsregister\*"; DestDir: "{app}\VB98\Wizards"; Flags: restartreplace uninsneveruninstall overwritereadonly recursesubdirs regserver noregerror; Tasks: licensefull 
Source: "Wizardsregtlib\*"; DestDir: "{app}\VB98\Wizards"; Flags: restartreplace uninsneveruninstall overwritereadonly recursesubdirs regtypelib noregerror; Tasks: licensefull 

;VB6Porter VB6Namespaces VB65 VBAPorter 
Source: "UpdateVB6Resource.exe"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly
Source: "RegAsm.exe"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly; Tasks: licenseVB6Porter
Source: "VBANET.tlb"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly regtypelib; Tasks: licenseVB6Porter
Source: "VBANET.dll"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly; Tasks: licenseVB6Porter
Source: "VBANET.pdb"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly; Tasks: licenseVB6Porter
Source: "VB6Porter.dll"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly regserver; Tasks: licenseVB6Porter
Source: "VB6Porter.exp"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly; Tasks: licenseVB6Porter
Source: "VB6Porter.lib"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly; Tasks: licenseVB6Porter
Source: "VBAPorter.dll"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly regserver; Tasks: licenseVB6Porter
Source: "VBAPorter.exp"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly; Tasks: licenseVB6Porter
Source: "VBAPorter.lib"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly; Tasks: licenseVB6Porter
Source: "VBANET.manifest"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace uninsneveruninstall overwritereadonly; Tasks: licenseVB6Porter
Source: "EULA.txt"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly
Source: "VSPROEULA.txt"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly; Tasks: licenseFull\licensePro  
Source: "VSENTEULA.txt"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly; Tasks: licenseFull\licenseEnt
Source: "VBA65EULA.txt"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly; Tasks: licensefull
Source: "SP6ENTEULA.txt"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly; Tasks: licenseFull\licenseEnt 
Source: "SP6PROEULA.txt"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly; Tasks: licenseFull\licensePro
Source: "vbporter64.ico"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly; Tasks: licenseVB6Porter
Source: "vbpproj64.ico"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly; Tasks: licenseVB6Porter
Source: "vbpClass64.ico"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly; Tasks: licenseVB6Porter
Source: "vbpBaseModule64.ico"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly; Tasks: licenseVB6Porter
Source: "vbpForm64.ico"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly; Tasks: licenseVB6Porter
Source: "vbpUC64.ico"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly; Tasks: licenseVB6Porter
Source: "vbpDesigner64.ico"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly; Tasks: licenseVB6Porter
Source: "vbpGroup64.ico"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly; Tasks: licenseVB6Porter
Source: "vbpWizard64.ico"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly; Tasks: licenseVB6Porter
Source: "vbpTLB.ico"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly; Tasks: licenseVB6Porter
Source: "vbpFRX.ico"; DestDir: {app}\VB98; Flags: ignoreversion restartreplace overwritereadonly; Tasks: licenseVB6Porter
Source: "VB60SP6-KB2708437-x86-ENU.msi"; DestDir: {tmp}; Flags: ignoreversion deleteafterinstall overwritereadonly; Tasks: licensefull
Source: "VB60SP6-KB3096896-x86-ENU.msi"; DestDir: {tmp}; Flags: ignoreversion deleteafterinstall overwritereadonly; Tasks: licensefull

[Dirs]
Name: "{app}\VB98\Projects"; Permissions: users-full; Flags: uninsneveruninstall; Tasks: licensefull

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\VB98\{#MyAppExeName1}"; IconFilename: "{app}\VB98\vbporter64.ico"; AfterInstall: SetElevationBit('{autoprograms}\VB6Porter.lnk'); Tasks: licensefull 
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\VB98\{#MyAppExeName1}"; IconFilename: "{app}\VB98\vbporter64.ico"; AfterInstall: SetElevationBit('{autodesktop}\VB6Porter.lnk'); Tasks: licensefull 
Name: "{group}\VB6Porter"; Filename: "{app}\VB98\VB6.EXE"; IconFilename: "{app}\VB98\vbporter64.ico"; WorkingDir: "{app}\VB98"; Tasks: licensefull 
Name: "{group}\Uninstall VB6Porter"; Filename: "{uninstallexe}"

[Code]
function InitializeUninstall(): Boolean;
  var ErrorCode: Integer;
begin
  ShellExec('open','taskkill.exe','/f /im {#MyAppExeName1}','',SW_HIDE,ewNoWait,ErrorCode); 
  ShellExec('open','taskkill.exe','/f /im {#MyAppExeName2}','',SW_HIDE,ewNoWait,ErrorCode); 
  ShellExec('open','taskkill.exe','/f /im {#MyAppExeName3}','',SW_HIDE,ewNoWait,ErrorCode); 
  ;ShellExec('open','tskill.exe',' {#MyAppExeName1}','',SW_HIDE,ewNoWait,ErrorCode);
  result := True;
end;

procedure SetElevationBit(Filename: string);
var
  Buffer: string;
  Stream: TStream;
begin
  Filename := ExpandConstant(Filename);
  Log('Setting elevation bit for ' + Filename);

  Stream := TFileStream.Create(FileName, fmOpenReadWrite);
  try
    Stream.Seek(21, soFromBeginning);
    SetLength(Buffer, 1);
    Stream.ReadBuffer(Buffer, 1);
    Buffer[1] := Chr(Ord(Buffer[1]) or $20);
    Stream.Seek(-1, soFromCurrent);
    Stream.WriteBuffer(Buffer, 1);
  finally
    Stream.Free;
  end;
end;

var
  CustomQueryPage: TInputQueryWizardPage;
procedure AddCustomQueryPage();
begin
  CustomQueryPage := CreateInputQueryPage(
    wpWelcome,
    'Product key identification',
    'This is the serial number that appears in the about dialog box inside of the VB6 IDE.',
    'Please enter your 10 digit product key, ie 123-1234567');

  { Add items (False means it's not a password edit) }
  CustomQueryPage.Add('Product key/serial:', False);
end;

procedure InitializeWizard();
begin
  AddCustomQueryPage();
end;

var
  ResultCode: Integer;
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
  if not Exec(ExpandConstant('{app}\VB98\UpdateVB6Resource.exe'), CustomQueryPage.Values[0], '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
   begin
    MsgBox('Failed to install product id' + #13#10 + SysErrorMessage(ResultCode), mbError, MB_OK);
  end;
  end;
end;

function CmdLineParamExists(const Value: string): Boolean;
var
  I: Integer;  
begin
  Result := False;
  for I := 1 to ParamCount do
    if CompareText(ParamStr(I), Value) = 0 then
    begin
      Result := True;
      Exit;
    end;
end;

procedure InstallProgramB;
var
ResultCode: Integer;
begin
  // Install programB and wait for it to terminate    Exec('reg.exe', 'import C:\Support\Banners.reg', '', SW_HIDE, ewWaitUntilTerminated, Code);
  if not Exec('reg.exe', 'import ' + expandconstant('{sys}\lic.reg') + '', '', SW_SHOWNORMAL, ewWaitUntilTerminated, ResultCode) then
   begin
    MsgBox('Failed to install programB!' + #13#10 + SysErrorMessage(ResultCode), mbError, MB_OK);
  end;
end;

procedure InstallProgramC;
var
ResultCode: Integer;
begin
    if not Exec('reg.exe', 'import ' + expandconstant('{sys}\customUI.reg') + '', '', SW_SHOWNORMAL, ewWaitUntilTerminated, ResultCode) then
   begin
    MsgBox('Failed to install programc!' + #13#10 + SysErrorMessage(ResultCode), mbError, MB_OK);
  end;
end;

[Run] 
;Filename: "msiexec.exe"; StatusMsg: "Installing VBA 65 SDK..."; Parameters: "/i ""{app}\VB98\vba65\setup.exe"" /passive"; Flags:runhidden runascurrentuser skipifsilent; Tasks: licensefull 
Filename: "{app}\VB98\regasm.exe"; Parameters: /tlb VBANET.dll; StatusMsg: "Installing VBANET.dll..."; Flags:runhidden runascurrentuser skipifsilent; Tasks: licensefull  
Filename: "{app}\VB98\regasm.exe"; Parameters: /codebase VBANET.tlb; StatusMsg: "Installing VBANET.tlb..."; Flags:runhidden runascurrentuser skipifsilent; Tasks: licensefull 
Filename: "msiexec.exe"; StatusMsg: "Installing Cumulative Update for Microsoft Visual Basic 6.0 SP6 (KB926857)..."; Parameters: "/i ""{tmp}\VB60SP6-KB926857-x86-ENU.msi"" /quiet"; Tasks: licensefull 
Filename: "msiexec.exe"; StatusMsg: "Installing Cumulative Update for Microsoft Visual Basic 6.0 SP6 (KB2641426)..."; Parameters: "/i ""{tmp}\VB60SP6-KB2641426-x86-ENU.msi"" /quiet"; Tasks: licensefull 
Filename: "msiexec.exe"; StatusMsg: "Installing Cumulative Update for Microsoft Visual Basic 6.0 SP6 (KB2708437)...."; Parameters: "/i ""{tmp}\VB60SP6-KB2708437-x86-ENU.msi"" /quiet"; Tasks: licensefull 
Filename: "msiexec.exe"; StatusMsg: "Installing Cumulative Update for Microsoft Visual Basic 6.0 SP6 (KB3096896)..."; Parameters: "/i ""{tmp}\VB60SP6-KB3096896-x86-ENU.msi"" /quiet"; Tasks: licensefull 



[registry]
Root: HKCU; Subkey: "SOFTWARE\Microsoft\Visual Basic\6.0\Addins\VbaIW.Wizard"; ValueType: string; ValueName: "LoadBehavior"; ValueData: 0; Tasks: licensefull

Root: HKCU; Subkey: "SOFTWARE\Microsoft\Visual Basic\6.0"; ValueType: string; ValueName: "URLOwnersArea"; ValueData: ""; Tasks: licensefull
;core activex controls
Root: HKCR; Subkey: "Licenses\6000720D-F342-11D1-AF65-00A0C90DCA10"; ValueType: string; ValueName: ""; ValueData: "kefeflhlhlgenelerfleheietfmflelljeqf"; Tasks: licenseFull\licenseEnt 
;pro or enterprise license
Root: HKCR; Subkey: "Licenses\74872840-703A-11d1-A3AF-00A0C90F26FA"; ValueType: string; ValueName: ""; ValueData: ""; Tasks: licenseFull\licenseLearning 
Root: HKCR; Subkey: "Licenses\74872841-703A-11d1-A3AF-00A0C90F26FA"; ValueType: string; ValueName: ""; ValueData: ""; Tasks: licenseFull\licenseLearning   
Root: HKCR; Subkey: "Licenses\74872840-703A-11d1-A3AF-00A0C90F26FA"; ValueType: string; ValueName: ""; ValueData: "mninuglgknogtgjnthmnggjgsmrmgniglish"; Tasks: licenseFull\licensePro 
Root: HKCR; Subkey: "Licenses\74872841-703A-11d1-A3AF-00A0C90F26FA"; ValueType: string; ValueName: ""; ValueData: ""; Tasks: licenseFull\licensePro 
Root: HKCR; Subkey: "Licenses\74872840-703A-11d1-A3AF-00A0C90F26FA"; ValueType: string; ValueName: ""; ValueData: "mninuglgknogtgjnthmnggjgsmrmgniglish"; Tasks: licenseFull\licenseEnt 
Root: HKCR; Subkey: "Licenses\74872841-703A-11d1-A3AF-00A0C90F26FA"; ValueType: string; ValueName: ""; ValueData: "klglsejeilmereglrfkleeheqkpkelgejgqf"; Tasks: licenseFull\licenseEnt

;dao
Root: HKCR; Subkey: "Licenses\F4FC596D-DFFE-11CF-9551-00AA00A3DC45"; ValueType: string; ValueName: ""; ValueData: "mbmabptebkjcdlgtjmskjwtsdhjbmkmwtrak"; Tasks: licensefull

;ide code editor colors
Root: HKCU; Subkey: "SOFTWARE\Microsoft\VBA\Microsoft Visual Basic"; ValueType: string; ValueName: "CodeBackColors"; ValueData: "0 0 1 7 6 0 0 0 0 0 0 0 0 0 0 0 "; Flags: uninsdeletevalue
Root: HKCU; Subkey: "SOFTWARE\Microsoft\VBA\Microsoft Visual Basic"; ValueType: string; ValueName: "CodeForeColors"; ValueData: "5 0 3 0 1 10 13 4 0 0 0 0 0 0 0 0 "; Flags: uninsdeletevalue
Root: HKCU; Subkey: "SOFTWARE\Microsoft\VBA\6.0\Common"; ValueType: string; ValueName: "CodeBackColors"; ValueData: "0 0 1 7 6 0 0 0 0 0 0 0 0 0 0 0 "; Flags: uninsdeletevalue
Root: HKCU; Subkey: "SOFTWARE\Microsoft\VBA\6.0\Common"; ValueType: string; ValueName: "CodeForeColors"; ValueData: "5 0 3 0 1 10 13 4 0 0 0 0 0 0 0 0 "; Flags: uninsdeletevalue

Root: HKCU; Subkey: "SOFTWARE\Microsoft\VBA\Microsoft Visual Basic"; ValueType: dword; ValueName: "SyntaxChecking"; ValueData: "00000000"; Flags: uninsdeletevalue

;icon associations   dword:
Root: HKCR; Subkey: ".vbp"; ValueType: string; ValueName: ""; ValueData: "VB6ProjectFile"; Flags: uninsdeletevalue
Root: HKCR; Subkey: "VB6ProjectFile"; ValueType: string; ValueName: ""; ValueData: "Visual Basic Project"; Flags: uninsdeletekey
;Root: HKCR; Subkey: "VB6ProjectFile\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\VB98\VB6.EXE,16" 
Root: HKCR; Subkey: "VB6ProjectFile\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\VB98\vbpproj64.ico" 
Root: HKCR; Subkey: "VB6ProjectFile\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\VB98\VB6.EXE"" ""%1""" 

Root: HKCR; Subkey: ".cls"; ValueType: string; ValueName: ""; ValueData: "VB6ClassFile"; Flags: uninsdeletevalue
Root: HKCR; Subkey: "VB6ClassFile"; ValueType: string; ValueName: ""; ValueData: "Visual Basic Class Module"; Flags: uninsdeletekey  
;Root: HKCR; Subkey: "VB6ClassFile\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\VB98\VB6.EXE,9" 
Root: HKCR; Subkey: "VB6ClassFile\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\VB98\vbpClass64.ico"  
Root: HKCR; Subkey: "VB6ClassFile\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\VB98\VB6.EXE"" ""%1"""  

Root: HKCR; Subkey: ".bas"; ValueType: string; ValueName: ""; ValueData: "VB6ModuleFile"; Flags: uninsdeletevalue 
Root: HKCR; Subkey: "VB6ModuleFile"; ValueType: string; ValueName: ""; ValueData: "Visual Basic Module"; Flags: uninsdeletekey  
Root: HKCR; Subkey: "VB6ModuleFile\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\VB98\vbpBaseModule64.ico"
;Root: HKCR; Subkey: "VB6ModuleFile\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\VB98\VB6.EXE,8"
Root: HKCR; Subkey: "VB6ModuleFile\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\VB98\VB6.EXE"" ""%1""" 

Root: HKCR; Subkey: ".frm"; ValueType: string; ValueName: ""; ValueData: "VB6FormFile"; Flags: uninsdeletevalue
Root: HKCR; Subkey: "VB6FormFile"; ValueType: string; ValueName: ""; ValueData: "Visual Basic Form File"; Flags: uninsdeletekey 
;Root: HKCR; Subkey: "VB6FormFile\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\VB98\VB6.EXE,1"  
Root: HKCR; Subkey: "VB6FormFile\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\VB98\vbpForm64.ico" 
Root: HKCR; Subkey: "VB6FormFile\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\VB98\VB6.EXE"" ""%1"""
;
Root: HKCR; Subkey: ".ctl"; ValueType: string; ValueName: ""; ValueData: "VB6UCFile"; Flags: uninsdeletevalue  
Root: HKCR; Subkey: "VB6UCFile"; ValueType: string; ValueName: ""; ValueData: "Visual Basic User Control"; Flags: uninsdeletekey
;Root: HKCR; Subkey: "VB6UCFile\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\VB98\VB6.EXE,5"
Root: HKCR; Subkey: "VB6UCFile\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\VB98\vbpUC64.ico"
Root: HKCR; Subkey: "VB6UCFile\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\VB98\VB6.EXE"" ""%1"""
;
Root: HKCR; Subkey: ".dsr"; ValueType: string; ValueName: ""; ValueData: "VB6DesFile"; Flags: uninsdeletevalue 
Root: HKCR; Subkey: "VB6DesFile"; ValueType: string; ValueName: ""; ValueData: "Visual Basic Designer Module"; Flags: uninsdeletekey  
;Root: HKCR; Subkey: "VB6DesFile\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\VB98\VB6.EXE,4"  
Root: HKCR; Subkey: "VB6DesFile\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\VB98\vbpDesigner64.ico"  
Root: HKCR; Subkey: "VB6DesFile\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\VB98\VB6.EXE"" ""%1""" 
;
Root: HKCR; Subkey: ".vbg"; ValueType: string; ValueName: ""; ValueData: "VB6PGFile"; Flags: uninsdeletevalue
Root: HKCR; Subkey: "VB6PGFile"; ValueType: string; ValueName: ""; ValueData: "Visual Basic Group Project"; Flags: uninsdeletekey  
;Root: HKCR; Subkey: "VB6PGFile\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\VB98\VB6.EXE,22" 
Root: HKCR; Subkey: "VB6PGFile\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\VB98\vbpGroup64.ico" 
Root: HKCR; Subkey: "VB6PGFile\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\VB98\VB6.EXE"" ""%1"""
;
Root: HKCR; Subkey: ".vbz"; ValueType: string; ValueName: ""; ValueData: "VB6ZFile"; Flags: uninsdeletevalue 
Root: HKCR; Subkey: "VB6ZFile"; ValueType: string; ValueName: ""; ValueData: "Visual Basic Wizard"; Flags: uninsdeletekey
;Root: HKCR; Subkey: "VB6ZFile\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\VB98\VB6.EXE,21"  
Root: HKCR; Subkey: "VB6ZFile\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\VB98\vbpWizard64.ico"  
Root: HKCR; Subkey: "VB6ZFile\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\VB98\VB6.EXE"" ""%1""" 
 ;
Root: HKCR; Subkey: ".frx"; ValueType: string; ValueName: ""; ValueData: "VB6frxFile"; Flags: uninsdeletevalue 
Root: HKCR; Subkey: "VB6frxFile"; ValueType: string; ValueName: ""; ValueData: "Visual FoxPro Report"; Flags: uninsdeletekey
Root: HKCR; Subkey: "VB6frxFile\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\VB98\vbpFRX.ico"  
Root: HKCR; Subkey: "VB6frxFile\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\VB98\VB6.EXE"" ""%1""" 
  ;
Root: HKCR; Subkey: ".tlb"; ValueType: string; ValueName: ""; ValueData: "VB6TLBFile"; Flags: uninsdeletevalue 
Root: HKCR; Subkey: "VB6TLBFile"; ValueType: string; ValueName: ""; ValueData: "Type Library"; Flags: uninsdeletekey
Root: HKCR; Subkey: "VB6TLBFile\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\VB98\vbpTLB.ico"  
Root: HKCR; Subkey: "VB6TLBFile\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\VB98\VB6.EXE"" ""%1""" 


[UninstallRun]
Filename: "{sys}\regsvr32.exe"; Parameters: "/u /s VB6Porter.dll"; WorkingDir: "{app}\VB98"; Flags:runhidden runascurrentuser; Tasks: licenseVB6Porter  
Filename: "{sys}\regsvr32.exe"; Parameters: "/u /s VBAPorter.dll"; WorkingDir: "{app}\VB98"; Flags:runhidden runascurrentuser; Tasks: licenseVB6Porter  
Filename: "{app}\VB98\regasm.exe"; Parameters: "/unregister VBANET.dll"; Flags:runhidden runascurrentuser; Tasks: licenseVB6Porter  
                 
[UninstallDelete] 
