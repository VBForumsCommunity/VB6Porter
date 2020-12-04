VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAPCIntegration"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Const ERROR_OUT_OF_MEM     As Long = 0
Private Const ERROR_BAD_FORMAT     As Long = 11
Private Const ERROR_FILE_NOT_FOUND As Long = 2
Private Const ERROR_PATH_NOT_FOUND As Long = 3
Private Const ERROR_SUCCESS        As Long = 32
Private Const ERROR_NO_ASSOC       As Long = 31
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    stFilter As String
    stCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    strFile As String
    nMaxFile As Long
    stFileTitle As String
    nMaxFileTitle As Long
    stInitialDir As String
    strTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    stDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Type LineComment
    lcode As String
    lcomment As String
End Type
Public Enum procScope
    ScopePrivate = 1
    ScopePublic = 2
    ScopeFriend = 3
    ScopeDefault = 4
End Enum
Public Enum LineSplits
    LineSplitRemove = 0
    LineSplitKeep = 1
    LineSplitConvert = 2
End Enum
Public Type ProcInfo
    procName As String
    ProcKind As VBIDE.vbext_ProcKind
    ProcStartLine As Long
    ProcBodyLine As Long
    ProcCountLines As Long
    procScope As procScope
    ProcDeclaration As String
End Type
Private Enum OFN_FLAGS
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
End Enum

Private Declare Function apiGetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (ByRef ofn As OPENFILENAME) As Long
Private Declare Function apiGetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (ByRef ofn As OPENFILENAME) As Long
Private Declare Function apiWinExec Lib "kernel32" Alias "WinExec" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Private Declare Function apiSetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function apiSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long) As Long
Private Declare Function apiVkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar2 As Long) As Long
Private Const IDC_APPSTARTING As Long = 32650&
Private Const GCW_HCURSOR     As Long = (-12)
Private Const IDC_WAIT        As Long = 32514&
Private Const OCR_NORMAL      As Long = 32512
Private Declare Function apiLoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function apiSetClassWord Lib "user32" Alias "SetClassWord" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Private Declare Function apiDestroyCursor Lib "user32" Alias "DestroyCursor" (ByVal hCursor As Long) As Long
Private Declare Function apiSetSystemCursor Lib "user32" Alias "SetSystemCursor" (ByVal hcur As Long, ByVal id As Long) As Long
Dim SysCursHandle                                          As Long
Dim Curs1Handle                                            As Long
Private WithEvents m_apcHost                               As MSAPC.apc
Attribute m_apcHost.VB_VarHelpID = -1
Private WithEvents m_ApcProject                            As MSAPC.Project
Attribute m_ApcProject.VB_VarHelpID = -1
Private WithEvents m_apcReferenceEvents                    As VBIDE.ReferencesEvents
Attribute m_apcReferenceEvents.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventNewProject          As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventNewProject.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventOpenProject         As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventOpenProject.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventAddProject          As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventAddProject.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventRemoveProject       As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventRemoveProject.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventPrint               As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventPrint.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventSave                As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventSave.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventSaveAs              As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventSaveAs.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventSaveMod             As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventSaveMod.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventSaveModAs           As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventSaveModAs.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventMake                As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventMake.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventMakeGroup           As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventMakeGroup.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventImport              As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventImport.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventExport              As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventExport.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventExit                As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventExit.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventAddFile             As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventAddFile.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventReferences          As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventReferences.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventComponents          As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventComponents.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventAddControls         As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventAddControls.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventProperties          As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventProperties.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventAddProcedure        As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventAddProcedure.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventOptions             As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventOptions.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventRun                 As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventRun.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventRun2                As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventRun2.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventBreak2              As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventBreak2.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventEnd2                As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventEnd2.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventStart               As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventStart.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventBreak               As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventBreak.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventEnd                 As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventEnd.VB_VarHelpID = -1
'
Private WithEvents m_apcCommandBarEventToolbox             As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventToolbox.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventObjectBrowser       As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventObjectBrowser.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventFind                As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventFind.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventUndo                As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventUndo.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventRedo                As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventRedo.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventRemove              As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventRemove.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventComment             As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventComment.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventUnComment           As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventUnComment.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventSave2               As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventSave2.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventMake2               As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventMake2.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventStart2              As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventStart2.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventFormat              As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventFormat.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventMetric              As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventMetric.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventCodeFormat          As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventCodeFormat.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventBack                As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventBack.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventCodeToggleExpansion As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventCodeToggleExpansion.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventCodeCollapseAll     As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventCodeCollapseAll.VB_VarHelpID = -1
Private WithEvents m_apcCommandBarEventSyntaxCheck         As VBIDE.CommandBarEvents
Attribute m_apcCommandBarEventSyntaxCheck.VB_VarHelpID = -1
Private Sub Class_Initialize()
    Set m_apcHost = New apc
    ReDim Preserve pSettings(0)
    m_apcHost.ContinueOnError = True 'uncomment to make vba 6.0/6.1 compatible
End Sub
Private Sub Class_Terminate()
    On Error Resume Next
    Set vbEnv = Nothing
    Set m_apcHost = Nothing
    Set m_ApcProject = Nothing
    Set m_apcReferenceEvents = Nothing
    Set m_apcCommandBarEventNewProject = Nothing
    Set m_apcCommandBarEventOpenProject = Nothing
    Set m_apcCommandBarEventAddProject = Nothing
    Set m_apcCommandBarEventRemoveProject = Nothing
    Set m_apcCommandBarEventPrint = Nothing
    Set m_apcCommandBarEventSave = Nothing
    Set m_apcCommandBarEventSaveAs = Nothing
    Set m_apcCommandBarEventSaveMod = Nothing
    Set m_apcCommandBarEventSaveModAs = Nothing
    Set m_apcCommandBarEventMake = Nothing
    Set m_apcCommandBarEventMakeGroup = Nothing
    Set m_apcCommandBarEventImport = Nothing
    Set m_apcCommandBarEventExport = Nothing
    Set m_apcCommandBarEventExit = Nothing
    Set m_apcCommandBarEventAddFile = Nothing
    Set m_apcCommandBarEventReferences = Nothing
    Set m_apcCommandBarEventComponents = Nothing
    Set m_apcCommandBarEventAddControls = Nothing
    Set m_apcCommandBarEventProperties = Nothing
    Set m_apcCommandBarEventAddProcedure = Nothing
    Set m_apcCommandBarEventOptions = Nothing
    Set m_apcCommandBarEventRun = Nothing
    Set m_apcCommandBarEventStart = Nothing
    Set m_apcCommandBarEventBreak = Nothing
    Set m_apcCommandBarEventEnd = Nothing
    Set m_apcCommandBarEventRun2 = Nothing
    Set m_apcCommandBarEventBreak2 = Nothing
    Set m_apcCommandBarEventEnd2 = Nothing
    Set m_apcCommandBarEventRemove = Nothing
    '
    Set m_apcCommandBarEventToolbox = Nothing
    Set m_apcCommandBarEventObjectBrowser = Nothing
    Set m_apcCommandBarEventFind = Nothing
    Set m_apcCommandBarEventUndo = Nothing
    Set m_apcCommandBarEventRedo = Nothing
    Set m_apcCommandBarEventComment = Nothing
    Set m_apcCommandBarEventUnComment = Nothing
    Set m_apcCommandBarEventSave2 = Nothing
    Set m_apcCommandBarEventMake2 = Nothing
    Set m_apcCommandBarEventStart2 = Nothing
    Set m_apcCommandBarEventFormat = Nothing
    Set m_apcCommandBarEventMetric = Nothing
    Set m_apcCommandBarEventCodeFormat = Nothing
    Set m_apcCommandBarEventBack = Nothing
    Set m_apcCommandBarEventCodeToggleExpansion = Nothing
    Set m_apcCommandBarEventCodeCollapseAll = Nothing
    Set m_apcCommandBarEventSyntaxCheck = Nothing
    If hmbutton <> 0 Then apiUnhookWindowsHookEx hmbutton
    If hkey <> 0 Then apiUnhookWindowsHookEx hkey
End Sub
Friend Sub FormatCodeModule(ByVal cm As CodeModule, Optional ByVal auto As Boolean = False)
    On Error Resume Next
    If cfAutoReformat = False And auto = True Then Exit Sub
    Dim sl As Long
    Dim sc As Long
    Dim el As Long
    Dim ec As Long
    Call cm.CodePane.GetSelection(sl, sc, el, ec)
    Dim C     As String
    Dim lns() As String
    lns = Split(cm.Lines(1, cm.CountOfLines), vbCrLf)
    Dim ln1      As Variant
    Dim x        As Long
    Dim bl       As Long
    Dim lc       As Long
    Dim incase   As Boolean
    Dim oldpoint As Long
    For Each ln1 In lns
        lc = lc + 1
        Dim ln As String
        ln = LCase(Trim(ln1))
        If auto = True Or ln <> "" Or (ln = "" And cfReformatBlank = False) Then
            If ln = "loop" Or InStr(1, ln, "loop ") = 1 Or InStr(1, ln, "loop until ") = 1 Then x = x - 3
            If ln = "next" Or InStr(1, ln, "next ") = 1 Then x = x - 3
            If ln = "end with" Or InStr(1, ln, "end with") = 1 Then x = x - 3
            If ln = "wend" Or InStr(1, ln, "wend") = 1 Then x = x - 3
            If InStr(1, ln, "end if") = 1 Then x = x - 3
            If InStr(1, ln, "elseif ") = 1 Then x = x - 3
            If ln = "else" Or InStr(1, ln, "else ") = 1 Then x = x - 3
            If InStr(1, ln, "end select") = 1 Then x = x - 3
            If InStr(1, ln, "case ") = 1 Then x = x - 3
            If InStr(1, ln, "end type") = 1 Then x = 0
            If InStr(1, ln, "end enum") = 1 Then x = 0
            If InStr(1, ln, "end property") = 1 Then x = 0
            If InStr(1, ln, "type ") = 1 Then x = 0
            If InStr(1, ln, "private type ") = 1 Then x = 0
            If InStr(1, ln, "public type ") = 1 Then x = 0
            If InStr(1, ln, "enum ") = 1 Then x = 0
            If InStr(1, ln, "private property ") = 1 Then x = 0
            If InStr(1, ln, "public property ") = 1 Then x = 0
            If InStr(1, ln, "property ") = 1 Then x = 0
            If InStr(1, ln, "private enum ") = 1 Then x = 0
            If InStr(1, ln, "public enum ") = 1 Then x = 0
            If InStr(1, ln, "end sub") = 1 Then x = 0
            If InStr(1, ln, "end function") = 1 Then x = 0
            If InStr(1, ln, "private sub ") = 1 Then x = 0
            If InStr(1, ln, "public sub ") = 1 Then x = 0
            If InStr(1, ln, "friend sub ") = 1 Then x = 0
            If InStr(1, ln, "private function ") = 1 Then x = 0
            If InStr(1, ln, "public function ") = 1 Then x = 0
            If InStr(1, ln, "friend function ") = 1 Then x = 0
            If InStr(1, ln, "function ") = 1 Then x = 0
            If InStr(1, ln, "sub ") = 1 Then x = 0
            If IsGoToPoint(ln) = True Then oldpoint = x: x = 0
            If x < 0 Then x = 0
            C = C & Space(x) & Trim(ln1) & vbCrLf
            If IsGoToPoint(ln) = True Then x = oldpoint
            If ln = "do" Or InStr(1, ln, "do ") = 1 Then x = x + 3 'Or InStr(1, ln, "do while ") = 1
            If InStr(1, ln, "for ") = 1 Then x = x + 3
            If InStr(1, ln, "while ") = 1 Then x = x + 3
            If InStr(1, ln, "with ") = 1 Then x = x + 3
            If InStr(1, ln, "if ") = 1 Then
                If Right(ln, 5) = " then" Or InStr(1, reducespaces(ln), " then '") <> 0 Then
                    x = x + 3
                End If
            End If
            If InStr(1, ln, "elseif ") = 1 Then x = x + 3
            If ln = "else" Or InStr(1, ln, "else ") = 1 Then x = x + 3
            If InStr(1, ln, "select case ") = 1 Then x = x + 3
            If InStr(1, ln, "case ") = 1 Then x = x + 3
            If InStr(1, ln, "private sub ") = 1 Then x = x + 3
            If InStr(1, ln, "public sub ") = 1 Then x = x + 3
            If InStr(1, ln, "friend sub ") = 1 Then x = x + 3
            If InStr(1, ln, "private function ") = 1 Then x = x + 3
            If InStr(1, ln, "public function ") = 1 Then x = x + 3
            If InStr(1, ln, "friend function ") = 1 Then x = x + 3
            If InStr(1, ln, "function ") = 1 Then x = x + 3
            If InStr(1, ln, "sub ") = 1 Then x = x + 3
            If InStr(1, ln, "type ") = 1 Then x = x + 3
            If InStr(1, ln, "private type ") = 1 Then x = x + 3
            If InStr(1, ln, "public type ") = 1 Then x = x + 3
            If InStr(1, ln, "enum ") = 1 Then x = x + 3
            If InStr(1, ln, "private enum ") = 1 Then x = x + 3
            If InStr(1, ln, "public enum ") = 1 Then x = x + 3
            If InStr(1, ln, "private property ") = 1 Then x = x + 3
            If InStr(1, ln, "public property ") = 1 Then x = x + 3
            If InStr(1, ln, "property ") = 1 Then x = x + 3
        Else
            If auto = False And lc < sl Then bl = bl + 1
        End If
    Next
    If Trim(C) <> "" And Trim(C) <> vbCrLf Then
        Call cm.DeleteLines(1, cm.CountOfLines)
        cm.InsertLines 1, C
        Call cm.CodePane.SetSelection(sl - bl, sc, sl - bl, sc)
        If Trim(cm.Lines(cm.CountOfLines, 1)) = "" Then cm.DeleteLines cm.CountOfLines, 1 'trim extra line inserted
    End If
    Dim l As String
    If cfReformatBlank = True And auto = False Then
        For lc = 1 To cm.CountOfLines
            l = cm.Lines(lc, 1)
            If Trim(l) = "" Then cm.DeleteLines lc, 1
        Next
    End If
    DoEvents
End Sub

Private Sub CutomizeIDE() 'Customize and remove menu items and windows to emulate VB6

    On Error Resume Next
    Dim indx As Long
    Dim n    As Long
    Dim i    As Long
    Dim cap  As String
    Dim cb   As Object
    Dim co2  As Object
    Dim co   As Object
    Dim z    As Long
    Dim cap2 As String
    '    Dim txt  As String
    Dim j    As Long
    For i = 1 To vbEnv.CommandBars.Count
        Set cb = vbEnv.CommandBars.Item(i)
        Dim cbName As String
        cbName = cb.Name
        'txt = txt & "Main: " & cbName & vbCrLf
        indx = vbEnv.CommandBars.Item(i).Index
        For n = 1 To vbEnv.CommandBars(indx).Controls.Count
            Set co2 = vbEnv.CommandBars(indx).Controls.Item(n)
            cap = co2.caption
            If cbName = "Standard" Then co2.visible = False
            If cbName = "Project Window" Then
                ' co2.visible = True
                If cap = "&Remove" Or cap = "&Import File..." Or cap = "&Export File..." Then co2.visible = False
                'MsgBox co2.caption
            End If
            '             If cap = "&File" Then 'If co2.Type = msoControlPopup Then
            '            For j = 1 To co2.Controls.Count
            '             Set co = co2.Controls(j)
            '             cap2 = co.enabled = False
            '
            '            Next
            '            End If
            '              If cap = "&File" Then
            '              co2.enabled = True
            ''                For j = 1 To co2.Count
            ''                   co2.Item(z).enabled = False
            ''                Next
            '              End If
            '            If InStr(1, cap, "&Remove ") <> 0 Then
            '               MsgBox cbName
            '            End If
            ' If Trim(cap) <> "" And Trim(cap) = "&Definition" Then
            '                txt = txt & "   Secondary: " & vbEnv.CommandBars(indx).Name & vbCrLf
            '                txt = txt & "    TooltipText: " & co2.ToolTipText & vbCrLf
            '                txt = txt & "    FaceId: " & co2.FaceId & vbCrLf
            '                txt = txt & "    DescriptionText: " & co2.DescriptionText & vbCrLf
            '                txt = txt & "    ID: " & co2.ID & vbCrLf
            '                txt = txt & vbCrLf
            ' If co2.Type = msoControlPopup Then
            '' For z = 1 To co2.Controls.Count
            ' Set co = co2.Controls(z)
            ' cap2 = co.caption
            '                        If co.Type = msoControlPopup Then
            '
            '                           MsgBox cbName & " - " & cap & " - " & cap2
            '                        End If
            '                        txt = txt & "       Caption: " & cap2 & vbCrLf
            '                        txt = txt & "       TooltipText: " & co.ToolTipText & vbCrLf
            '                        txt = txt & "       FaceId: " & co.FaceId & vbCrLf
            '                        txt = txt & "       DescriptionText: " & co.DescriptionText & vbCrLf
            '                        txt = txt & "       ID: " & co.ID & vbCrLf
            '                        txt = txt & vbCrLf
            '                        Set co = Nothing
            ''Next
            'End If
            ' End If
            Set co2 = Nothing
        Next
        Set cb = Nothing
    Next
    '    Clipboard.Clear
    '    Clipboard.SetText txt, 1
    vbEnv.Windows("Project - No Open Projects").visible = True
    vbEnv.Windows("Properties").visible = True
    vbEnv.Windows("Immediate").visible = True
    vbEnv.Windows("UserForm").visible = True
    '
    vbEnv.CommandBars("Menu Bar").visible = True
    vbEnv.CommandBars("Standard").visible = True
    
    Dim pd As String
    

    If freshstart = True Then
       vbEnv.CommandBars("UserForm").visible = False
       vbEnv.CommandBars("Edit").visible = False
       vbEnv.CommandBars("Debug").visible = False
    End If

    
    '
    vbEnv.CommandBars("Menu Bar").Controls("&File").Controls(3).visible = False
    vbEnv.CommandBars("Menu Bar").Controls("&File").Controls("&Import File...").visible = False
    vbEnv.CommandBars("Menu Bar").Controls("&File").Controls("&Export File...").visible = False
    vbEnv.CommandBars("Menu Bar").Controls("&File").Controls("&Remove").visible = False
    vbEnv.CommandBars("Menu Bar").Controls("&File").Controls("&Print...").visible = False
    vbEnv.CommandBars("Menu Bar").Controls("&File").Controls("&Close and Return to Visual Basic 6.5").visible = False
    '
    vbEnv.CommandBars("Menu Bar").Controls("&Insert").caption = "&Project"
    vbEnv.CommandBars("Menu Bar").Controls("&Project").Controls("&Procedure...").visible = False
    vbEnv.CommandBars("Menu Bar").Controls("&Project").Controls("&UserForm").caption = "Add &UserForm"
    vbEnv.CommandBars("Menu Bar").Controls("&Project").Controls("&Module").caption = "Add &Module"
    vbEnv.CommandBars("Menu Bar").Controls("&Project").Controls("&Class Module").caption = "Add &Class Module"
    vbEnv.CommandBars("Menu Bar").Controls("&Project").Controls("Fi&le...").visible = False
    '
    vbEnv.CommandBars("Menu Bar").Controls("&Debug").Controls("Compi&le ").caption = "Pre-Compi&le"
    '
    vbEnv.CommandBars("Menu Bar").Controls("&Run").Controls("Run Macro").visible = False
    vbEnv.CommandBars("Menu Bar").Controls("&Run").Controls("Brea&k").visible = False
    vbEnv.CommandBars("Menu Bar").Controls("&Run").Controls("&Reset").visible = False
    vbEnv.CommandBars("Menu Bar").Controls("&Run").Controls("Design &Mode").visible = False
    '
    vbEnv.CommandBars("Menu Bar").Controls("&Tools").Controls("&References...").visible = False
    vbEnv.CommandBars("Menu Bar").Controls("&Tools").Controls("&Additional Controls...").visible = False
    vbEnv.CommandBars("Menu Bar").Controls("&Tools").Controls("&Macros...").visible = False
    vbEnv.CommandBars("Menu Bar").Controls("&Tools").Controls("&Options...").visible = False
    vbEnv.CommandBars("Menu Bar").Controls("&Tools").Controls("Prop&erties...").visible = False
    vbEnv.CommandBars("Menu Bar").Controls("&Tools").Controls("&Digital Signature...").visible = False
    'Caption:
    vbEnv.CommandBars("Project Window").Controls("Prop&erties...").visible = False
    vbEnv.CommandBars("Project Window (Break)").Controls("Prop&erties...").visible = False
    '    Dim vbComBarButton As Office.CommandBarButton
    Dim vbComBarButton As Object
    Call InsertNewMenuItem(vbComBarButton, "File", "&New Project", 1, "Create a new Visual Basic project file", True, True, True) 'Create and insert new menu item
    Set m_apcCommandBarEventNewProject = vbEnv.Events.CommandBarEvents(vbComBarButton) 'set the event
    Call InsertNewMenuItem(vbComBarButton, "File", "&Open Project...", 23, "Open a Visual Basic project file", True, True, False)
    Set m_apcCommandBarEventOpenProject = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "File", "A&dd Project...", 1, "Add a Visual Basic project file", True, True, True)
    Set m_apcCommandBarEventAddProject = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "File", "&Remove Project", 1, "Remove a Visual Basic project file", True, True, False)
    Set m_apcCommandBarEventRemoveProject = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "File", "Sa&ve Project", 3, "Saves the active project", True, True, True)
    Set m_apcCommandBarEventSave = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "File", "Sav&e Project As...", 1, "Save project to path", True, True, False)
    Set m_apcCommandBarEventSaveAs = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "File", "&Save Code Module", 1, "Save the active code module", True, True, True)
    Set m_apcCommandBarEventSaveMod = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "File", "Save Code Module &As...", 1, "Save the active code module as file", True, True, False)
    Set m_apcCommandBarEventSaveModAs = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "File", "Save Se&lection", 1, "Save the selection", True, False, False)
    Call InsertNewMenuItem(vbComBarButton, "File", "Save C&hange Script", 1, "", True, False, False)
    Call InsertNewMenuItem(vbComBarButton, "File", "&Print Code Module...", 4, "", True, True, True)
    Set m_apcCommandBarEventPrint = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "File", "Print Set&up...", 511, "Print setup", True, False, False)
    Call InsertNewMenuItem(vbComBarButton, "File", "Ma&ke .exe...", 1771, "Makes a native executable file from the active Visual Basic project", True, True, True)
    Set m_apcCommandBarEventMake = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "File", "Make Project &Group... ", 746, "Makes a group project file", True, True, False)
   ' vbComBarButton.Picture = LoadPicture(app.Path & "\images\ProjectGroup16.bmp")
     vbComBarButton.Picture = frmVB6x.ImageList1.ListImages.Item(2).Picture   'LoadPicture(app.Path & "\images\ProjectGroup16.bmp")
    
    Set m_apcCommandBarEventMakeGroup = vbEnv.Events.CommandBarEvents(vbComBarButton)
    '
    Call InsertNewMenuItem(vbComBarButton, "File", "&Import File", 524, "", True, True, True)
    Set m_apcCommandBarEventImport = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "File", "&Export File", 525, "", True, True, False)
    Set m_apcCommandBarEventExport = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "File", "E&xit", 1, "Exit the Visual Basic environment", True, True, True)
    Set m_apcCommandBarEventExit = vbEnv.Events.CommandBarEvents(vbComBarButton)
    '_______________________
    Call InsertNewMenuItem(vbComBarButton, "Insert", "Add &User Control", 1, "", True, False, False)
    Call InsertNewMenuItem(vbComBarButton, "Insert", "Add &Property Page", 1, "", True, False, False)
    Call InsertNewMenuItem(vbComBarButton, "Insert", "Add User &Document", 1, "", True, False, False)
    Call InsertNewMenuItem(vbComBarButton, "Insert", "&Add File...", 1, "", True, True, False)
    Set m_apcCommandBarEventAddFile = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Insert", "&Remove Code Module", 1, "", True, True, True)
    Set m_apcCommandBarEventRemove = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Insert", "Refere&nces...", 942, "", True, True, True)
    Set m_apcCommandBarEventReferences = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Insert", "C&omponents... ", 1, "", True, True, False)
    Set m_apcCommandBarEventComponents = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Insert", "&Additional Controls...", 642, "", True, True, False)
    Set m_apcCommandBarEventAddControls = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Insert", "Prop&erties...", 2578, "", True, True, True)
    Set m_apcCommandBarEventProperties = vbEnv.Events.CommandBarEvents(vbComBarButton)
    '______________________
    Call InsertNewMenuItem(vbComBarButton, "Run", "&Start", 186, "Debug", True, True, False)
    Set m_apcCommandBarEventRun2 = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Run", "Start With &Full Compile", 5415, "", True, True, False)
    Set m_apcCommandBarEventStart = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Run", "Brea&k ", 189, "", True, True, False)
    Set m_apcCommandBarEventBreak2 = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Run", "&End", 228, "", True, True, False)
    Set m_apcCommandBarEventEnd2 = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Run", "&Restart", 1, "", True, False, False)
    '_______________________
    Call InsertNewMenuItem(vbComBarButton, "Tools", "Add &Procedure...", 559, "", True, True, False)
    Set m_apcCommandBarEventAddProcedure = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Tools", "Procudure &Attributes...", 1, "", True, False, False)
    Call InsertNewMenuItem(vbComBarButton, "Tools", "&Menu Editor...", 1, "", True, False, True)
    Call InsertNewMenuItem(vbComBarButton, "Tools", "&Options... ", 1, "", True, True, True)
    Set m_apcCommandBarEventOptions = vbEnv.Events.CommandBarEvents(vbComBarButton)
    '_______________________
    Call InsertNewMenuItem(vbComBarButton, "Standard", "Toolbo&x", 548, "", True, True, True)
    Set m_apcCommandBarEventToolbox = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Standard", "&Backport", 1759, "Backport VB.NET code to VB6 code", True, True, False)
    Set m_apcCommandBarEventBack = vbEnv.Events.CommandBarEvents(vbComBarButton) '[HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0\RecentFiles] 'Or VBA\6.0\RecentFiles
    Call InsertNewMenuItem(vbComBarButton, "Standard", "Code Formatting Options", 1, "", True, True, True)
    Set m_apcCommandBarEventCodeFormat = vbEnv.Events.CommandBarEvents(vbComBarButton)
   ' vbComBarButton.Picture = LoadPicture(app.Path & "\images\codesettings.bmp")
    vbComBarButton.Picture = frmVB6x.ImageList1.ListImages.Item(4).Picture
    
    Call InsertNewMenuItem(vbComBarButton, "Standard", "Code Metric Analyzer", 1, "", True, True, True)
    Set m_apcCommandBarEventMetric = vbEnv.Events.CommandBarEvents(vbComBarButton)
    'vbComBarButton.Picture = LoadPicture(app.Path & "\images\CodeMetric.bmp")
     vbComBarButton.Picture = frmVB6x.ImageList1.ListImages.Item(3).Picture
    
    Call InsertNewMenuItem(vbComBarButton, "Standard", "&Object Browser", 473, "", True, True, True)
    Set m_apcCommandBarEventObjectBrowser = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Standard", "&Find...", 141, "", True, True, True)
    Set m_apcCommandBarEventFind = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Standard", "Undo", 128, "", True, True, True)
    Set m_apcCommandBarEventUndo = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Standard", "Redo", 129, "", True, True, True)
    Set m_apcCommandBarEventRedo = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Standard", "Comment", 192, "", True, True, True)
    Set m_apcCommandBarEventComment = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Standard", "Uncomment", 2552, "", True, True, True)
    Set m_apcCommandBarEventUnComment = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Standard", "Reformat Code", 15, "Pretty listing (reformatting) of code", True, True, True)
    Set m_apcCommandBarEventFormat = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Standard", "Toggle Definitions", 15, "Toggles the expansion of all code procedures", True, True, True)
    'vbComBarButton.Picture = LoadPicture(app.Path & "\images\ToggleExpansion.bmp")
    vbComBarButton.Picture = frmVB6x.ImageList1.ListImages.Item(5).Picture
    
    Set m_apcCommandBarEventCodeCollapseAll = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Standard", "&Start", 186, "", True, True, True)
    Set m_apcCommandBarEventRun = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Standard", "Brea&k", 189, "", True, True, True)
    Set m_apcCommandBarEventBreak = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Standard", "Stop", 228, "", True, True, True)
    Set m_apcCommandBarEventEnd = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Standard", "Sa&ve Projects", 749, "Save projects", True, True, True)
    Set m_apcCommandBarEventSave2 = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Standard", "Make .exe...", 1771, "", True, True, True)
    Set m_apcCommandBarEventMake2 = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Standard", "Start With Full Compile", 5415, "Pre-compile, make, and start an executable file", True, True, True)
    Set m_apcCommandBarEventStart2 = vbEnv.Events.CommandBarEvents(vbComBarButton)
    Call InsertNewMenuItem(vbComBarButton, "Code Window", "Toggle Procedure &Expansion", 15, "Toggles the expansion of  the selected code procedure", True, True, True, 1)
   ' vbComBarButton.Picture = LoadPicture(app.Path & "\images\ToggleExpansion.bmp")
    vbComBarButton.Picture = frmVB6x.ImageList1.ListImages.Item(5).Picture
    
    Set m_apcCommandBarEventCodeToggleExpansion = vbEnv.Events.CommandBarEvents(vbComBarButton)
    ' Call InsertNewMenuItem(vbComBarButton, "Menu Bar", "Syntax error:", 1, "Displays syntax errors from previous line", True, True, True)
    'vbComBarButton.caption = "yup"
    '
    Set vbComBarButton = vbEnv.CommandBars("Menu Bar").Controls.Add(msoControlPopup)
    'Dim b As Office.CommandBarButton
    With vbComBarButton
        .caption = "Syntax Check:"
        .FaceId = 1
        .OnAction = "qtrReport"
        .BeginGroup = True
        .enabled = True
        .DescriptionText = "Auto Syntax Check"
        .visible = True
        .ToolTipText = "Auto Syntax Check"
        .Tag = "Auto Syntax Check"
    End With
    Set m_apcCommandBarEventSyntaxCheck = vbEnv.Events.CommandBarEvents(vbComBarButton)
    vbEnv.CommandBars("Standard").Controls("Undo").enabled = vbEnv.CommandBars.FindControl(msoControlButton, 128).enabled
    vbEnv.CommandBars("Standard").Controls("Redo").enabled = vbEnv.CommandBars.FindControl(msoControlButton, 129).enabled
    
    
'    vbEnv.ActiveVBProject.References.AddFromFile "C:\Program Files\Common Files\microsoft shared\VBA\VBA7\VBE7.DLL\3" '




End Sub




Private Sub InsertNewMenuItem(ByRef vbComBarButton As Object, ByVal parent As String, ByVal caption As String, ByVal FaceId As Long, ByVal desc As String, ByVal visible As Boolean, ByVal enabled As Boolean, ByVal separator As Boolean, Optional ByVal Index As Long = -1, Optional ByVal tg As String = vbNullString)
    On Error Resume Next
    Set vbComBarButton = vbEnv.CommandBars(parent).Controls.Add(msoControlButton)
    With vbComBarButton
        .caption = caption
        .FaceId = FaceId
        .OnAction = "qtrReport"
        .BeginGroup = separator
        .enabled = enabled
        .DescriptionText = desc
        .visible = visible
        .ToolTipText = desc
        .Tag = tg
    End With
End Sub
Friend Sub Initialize()
    On Error Resume Next
    m_apcHost.hwnd = frmVB6x.hwnd
    m_apcHost.ApplicationObject = New clsAPCApplication  ' This ought to be a set statement here, not a directassignment, but the Set will fail.
    m_apcHost.HostName = app.title
    m_apcHost.LicenseKey = "122333444455555666666777777788888888999999999"          '16175148714896599659AFABD8ED3C2A416B45E4CD6F5484BD8CED
    m_apcHost.RegistryKey = "vb6x" '[HKEY_CURRENT_USER\Software\Microsoft\VBA\6.0\vb6x]' The registry key where IDE customizations will be stored
    m_apcHost.Locale = 1033 'English
    m_apcHost.FileFilter = "All Reference Files (*.olb, *.tlb, *.dll, *.exe, *.ocx)" & Chr(0) & "*.olb;*.tlb;*.dll;*.exe;*.ocx" & Chr(0) ' File filter for the VBA references dialog
    m_apcHost.BackColor = RGB(255, 0, 255)
    Set vbEnv = m_apcHost.VBE
    m_ApcProject.Dirty = False
    isDirty = True
    CutomizeIDE 'customize to emulate Visual Basic 6.0
End Sub
Friend Property Get ApcHost() As MSAPC.apc
    Set ApcHost = m_apcHost
End Property
Friend Property Get ApcProject() As MSAPC.Project
    Set ApcProject = m_ApcProject
End Property
Friend Sub Show()
    On Error Resume Next
    m_apcHost.Ide.visible = True
  
End Sub
'Custome events
Private Sub m_apcCommandBarEventNewProject_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    frmNewProject.Show
    
'
'apc.ApcProject.Save
'apc.ApcProject.Dirty = False
'    vbEnv.ActiveVBProject.BuildFileName = "chop.exe"
'    vbEnv.ActiveVBProject.MakeCompiledFile
    
    
    
End Sub

Private Sub m_apcCommandBarEventOpenProject_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    OpenLoadProject
End Sub
Private Sub m_apcCommandBarEventAddFile_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    ImportAddFile
End Sub
Private Sub m_apcCommandBarEventAddProject_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    CreateNewProject
End Sub
Private Sub m_apcCommandBarEventRemoveProject_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    If vbEnv.ActiveVBProject Is Nothing Then Exit Sub
    Dim prjName As String
    prjName = vbEnv.ActiveVBProject.Name
    Dim dres As VbMsgBoxResult
    dres = MsgBox("Would you like to remove the selected project named: " & prjName, vbYesNo, "Confirm")
    If dres = vbYes Then
        If apc.CancelDirtyCheck(prjName) = False Then
            Dim vbproj As MSAPC.Project
            For Each vbproj In m_apcHost.Projects
                If vbproj.Name = prjName Then
                    vbproj.Close
                    Exit For
                End If
            Next
        End If
    End If
End Sub
Private Sub m_apcCommandBarEventSave_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    SaveAllProjectFiles vbEnv.ActiveVBProject
End Sub
Private Sub m_apcCommandBarEventSaveAs_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    If vbEnv.ActiveVBProject Is Nothing Then Exit Sub
    Dim fold As String
    Dim filt As String
    Dim inp  As String
    Dim f    As String
    Dim fso  As Object
    Dim inp2 As String
    Dim i    As Long
    Dim z    As Long
    For i = 1 To UBound(pSettings)
        If vbEnv.ActiveVBProject.Name = pSettings(i).Name Then
            z = i
            Exit For
        End If
    Next
    inp = vbEnv.ActiveVBProject.Name
    inp2 = pSettings(z).Startup
    filt = "Visual Basic Project Files" & vbNullChar & ".vbp" & vbNullChar & vbNullChar
    f = FileOpenSave(0, filt, 1, "vbp", inp, "Save File", -1, False)
    Set fso = CreateObject("Scripting.FileSystemObject")
    inp = fso.GetBaseName(f)
    fold = fso.GetParentFolderName(f)
    Set fso = Nothing
    If Trim(inp) = "" Then Exit Sub
    vbEnv.ActiveVBProject.Name = inp
    pSettings(z).Startup = inp2
    pSettings(z).IconForm = inp2
    SaveAllProjectFiles vbEnv.ActiveVBProject
End Sub
Private Sub m_apcCommandBarEventSaveMod_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    savecodemodas vbEnv.SelectedVBComponent.Name, GetProjectDirectory(vbEnv.ActiveVBProject)
End Sub
Private Sub m_apcCommandBarEventSaveModAs_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    Dim typ   As String
    Dim typ2  As String
    Dim filt  As String
    Dim f     As String
    Dim fso   As Object
    Dim fname As String
    Dim fold  As String
    Dim inp   As String
    Select Case vbEnv.SelectedVBComponent.Type
        Case vbext_ct_ActiveXDesigner
            typ = "*.exe": typ2 = "Executable Files"
        Case vbext_ct_ClassModule
            typ = "*.cls": typ2 = "Class Files"
        Case vbext_ct_Document
            typ = "*.doc": typ2 = "Document Files"
        Case vbext_ct_MSForm
            typ = "*.frm": typ2 = "Form Files"
        Case vbext_ct_StdModule
            typ = "*.bas": typ2 = "Basic Files"
        Case Else
            typ = "*.*"
    End Select
    filt = typ2 & vbNullChar & typ & vbNullChar & vbNullChar
    inp = vbEnv.SelectedVBComponent.Name
    f = FileOpenSave(0, filt, 1, Replace(typ, "*", ""), inp, "Save File", -1, False)
    Set fso = CreateObject("Scripting.FileSystemObject")
    fname = fso.GetBaseName(f)
    fold = fso.GetParentFolderName(f)
    Set fso = Nothing
    savecodemodas fname, fold
End Sub
Private Sub savecodemodas(ByVal inp As String, ByVal pd As String)
    On Error Resume Next
    savingall = True
    Dim vbproj As VBIDE.VBProject
    Set vbproj = vbEnv.ActiveVBProject
    If vbproj Is Nothing Then Exit Sub
    Dim i          As Long
    Dim fso        As Object
    Dim sw         As Object
    Dim cmname     As String
    Dim cExtension As String
    Dim cLineCount As Long
    Dim cCode      As String
    Dim cHeader    As String
    Dim mpth       As String
    Dim vbcomp     As VBIDE.VBComponent
    Set vbcomp = vbEnv.SelectedVBComponent ' vbproj.VBComponents.Item(z)
    If inp <> vbNullString Then RenameComponent vbEnv.SelectedVBComponent.Name, inp
    If Dir(pd, vbDirectory) = "" Then MkDir pd 'make the directory if it does not already exist
    cmname = vbcomp.Name
    cExtension = GetFileExtension(vbcomp)
    cLineCount = vbcomp.CodeModule.CountOfLines
    If cLineCount > 0 Then cCode = vbcomp.CodeModule.Lines(1, cLineCount)
    If cExtension = ".frm" Then
        cHeader = GetDefaultFormHeader(cmname)
    ElseIf cExtension = ".bas" Then
        cHeader = GetDefaultModuleHeader(cmname)
    ElseIf cExtension = ".cls" Then
        cHeader = GetDefaultClassHeader(cmname)
    Else 'use the default VBA header, ie nothing at all
        cHeader = "" 'Left(cCode, lastAttr)
    End If
    If cHeader <> "" Then cCode = cHeader & vbCrLf & cCode
    mpth = pd & "\" & cmname & cExtension
    If Dir(mpth, vbNormal) <> "" Then Kill mpth
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set sw = fso.CreateTextFile(mpth, True)
    sw.WriteLine cCode
    sw.Close
    Set vbproj = Nothing
    Set vbcomp = Nothing
    Set sw = Nothing
    Set fso = Nothing
End Sub
Private Sub m_apcCommandBarEventPrint_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    If vbEnv.ActiveVBProject Is Nothing Then
        MsgBox "You need an active project to print anything"
        Exit Sub
    End If
    vbEnv.CommandBars.FindControl(msoControlButton, 4).Execute
End Sub
Private Sub m_apcCommandBarEventMake_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    If vbEnv.ActiveVBProject Is Nothing Then
        MsgBox "You need an active project to make an executable file"
        Exit Sub
    End If
    MakeExe False, vbEnv.ActiveVBProject
End Sub
Private Sub m_apcCommandBarEventMakeGroup_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    If vbEnv.ActiveVBProject Is Nothing Then
        MsgBox "You need an active project to make a project group"
        Exit Sub
    End If
    If vbEnv.VBProjects.Count < 2 Then
        MsgBox "You need at least two projects open to make a group"
        Exit Sub
    End If
    Dim filt As String
    Dim ofs  As String
    Dim fso  As Object
    filt = "Visual Basic Project Group" & vbNullChar & ".vbg" & vbNullChar & vbNullChar
    ofs = FileOpenSave(0, filt, 1, "vbg", lastgroup, "Save Project Group File", -1, False)
    If Trim(ofs) = "" Then Exit Sub
    Dim ots   As Object
    Dim i     As Long
    Dim pname As String
    lastgroup = ofs
    SaveProjects
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ots = fso.OpenTextFile(ofs, 2, 1)
    ots.WriteLine "VBGROUP 5.0"
    For i = 1 To vbEnv.VBProjects.Count
        pname = vbEnv.VBProjects.Item(i).Name
        If i = 1 Then
            ots.WriteLine "StartupProject=" & pname & "\" & pname & ".vbp"
        Else
            ots.WriteLine "Project=" & pname & "\" & pname & ".vbp"
        End If
    Next
    ots.Close
    Set ots = Nothing
    Set fso = Nothing
    For i = 1 To vbEnv.VBProjects.Count
        pname = vbEnv.VBProjects.Item(i).Name
        MakeExe False, vbEnv.VBProjects.Item(i)
        DoEvents
        apiSleep 400
    Next
End Sub
Private Sub m_apcCommandBarEventImport_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    ImportAddFile
End Sub
Friend Sub ImportAddFile(Optional ByVal vbproj As VBProject = Nothing, Optional ByVal f As String = vbNullString)
    On Error Resume Next
    If vbproj Is Nothing Then Set vbproj = vbEnv.ActiveVBProject
        If vbproj Is Nothing Then
        MsgBox "You need an active project to import components"
        Exit Sub
    End If
    'Dim f    As String
    Dim fso  As Object
    Dim filt As String
    filt = "Basic Files" & vbNullChar & "*.bas" & vbNullChar & "Class Files" & vbNullChar & "*.cls" & vbNullChar & "Form Files" & vbNullChar & "*.frm" & vbNullChar & "All Files" & vbNullChar & "*.*" & vbNullChar & vbNullChar
    If f = vbNullString Then f = FileOpenSave(0, filt, 2, "cls", "", "Add File", -1, True) '"Visual Basic files (*.bas)|(*.cls)|(*.txt)|(*.*)"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.fileexists(f) = False Then Set fso = Nothing: Exit Sub
    Dim fname As String
    Dim ext   As String
    fname = fso.GetBaseName(f)
    ext = fso.GetExtensionName(f)
    ext = LCase(ext)
    If ext = "bas" Then
        AddModuleToProject vbproj, vbext_ct_StdModule, fname
    ElseIf ext = "cls" Then
        AddModuleToProject vbproj, vbext_ct_ClassModule, fname
    ElseIf ext = "frm" Then
        AddModuleToProject vbproj, vbext_ct_MSForm, fname
        MsgBox "Unsupported extension type (." & ext & ") has been converted to MSForm Designer (.dsr)", vbInformation, "Warning"
    ElseIf ext = "dsr" Then
        AddModuleToProject vbproj, vbext_ct_MSForm, fname
    ElseIf ext = "ctl" Then
        AddModuleToProject vbproj, vbext_ct_MSForm, fname
        MsgBox "Unsupported extension type (." & ext & ") has been converted to MSForm Designer (.dsr)", vbInformation, "Warning"
    ElseIf ext = "pag" Then
        MsgBox "unsupported extension type "
        Exit Sub
    ElseIf ext = "res" Then
        MsgBox "unsupported extension type "
        Exit Sub
    ElseIf ext = "txt" Then
        AddModuleToProject vbproj, vbext_ct_StdModule, fname
    ElseIf True Then
        AddModuleToProject vbproj, vbext_ct_StdModule, fname
        MsgBox "unknown extension type "
        Exit Sub
    End If
    Dim line       As String
    Dim sourcefile As Object
    Set sourcefile = fso.OpenTextFile(f, 1)
    Dim nCode  As String
    Dim endatt As Long
    While Not sourcefile.AtEndOfStream ' while we are not finished reading through the file
        line = sourcefile.Readline
        If Trim(LCase(Left(line, 15))) = "option explicit" Then
           line = "'" & line
        End If
        If endatt = 1 And Left(Trim(line), 10) <> "Attribute " Then
            endatt = 3
        End If
        If Left(Trim(line), 10) = "Attribute " Then endatt = 1
        If endatt < 2 Then line = "'" & line
        
        nCode = nCode & line & vbCrLf
    Wend
    sourcefile.Close
    vbproj.VBComponents(fname).CodeModule.AddFromString nCode
    Set fso = Nothing
End Sub
Private Sub m_apcCommandBarEventExport_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    If vbEnv.ActiveVBProject Is Nothing Then
        MsgBox "You need an active project to export components"
        Exit Sub
    End If
    vbEnv.CommandBars.FindControl(msoControlButton, 525).Execute
End Sub
Private Sub m_apcCommandBarEventExit_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    If CancelDirtyCheck = True Then Exit Sub
    isclosing = True
    Set apc.ApcHost.ApplicationObject = Nothing
    apc.ApcHost.hwnd = 0
    apc.ApcHost.Ide.visible = False
End Sub
Private Sub m_apcCommandBarEventRemove_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    If vbEnv.ActiveVBProject Is Nothing Then
        MsgBox "You need an active project to remove a code module"
        Exit Sub
    End If
    Dim vbproj As VBIDE.VBProject
    Set vbproj = vbEnv.ActiveVBProject
    If vbproj Is Nothing Then Exit Sub
    Dim vbcomp As VBIDE.VBComponent
    Set vbcomp = vbEnv.SelectedVBComponent
    If vbcomp Is Nothing Then
        MsgBox "You need to select a code module"
        Exit Sub
    End If
    If Not vbcomp Is Nothing Then
        Dim cmname As String
        cmname = vbcomp.Name
        Dim prjName As String
        prjName = vbEnv.ActiveVBProject.Name
        If cmname <> "" Then
            Dim dres As VbMsgBoxResult
            dres = MsgBox("Would you like to remove the selected module " & vbCrLf & cmname & " from " & prjName & "?", vbYesNo, "Confirm")
            If dres = vbYes Then
                savingall = True
                vbproj.VBComponents.Remove vbcomp
            End If
        End If
    End If
    Set vbproj = Nothing
    Set vbcomp = Nothing
End Sub
Private Sub m_apcCommandBarEventReferences_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    If vbEnv.ActiveVBProject Is Nothing Then
        MsgBox "You need an active project to change references"
        Exit Sub
    End If
    vbEnv.CommandBars.FindControl(msoControlButton, 942).Execute
    Set m_apcReferenceEvents = vbEnv.Events.ReferencesEvents(vbEnv.ActiveVBProject)
End Sub
Private Sub m_apcCommandBarEventComponents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    If vbEnv.ActiveVBProject Is Nothing Then
        MsgBox "You need an active project to change components"
        Exit Sub
    End If
    '  vbEnv.CommandBars("Menu Bar").Controls("&Project").Controls("C&omponents...").Execute
    frmComponents.Show vbModal
End Sub
Private Sub m_apcCommandBarEventAddControls_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    If vbEnv.ActiveVBProject Is Nothing Then
        MsgBox "You need an active project to change controls"
        Exit Sub
    End If
    If vbEnv.SelectedVBComponent.Type <> vbext_ct_MSForm Then
        MsgBox "Select Form to add controls"
        Exit Sub
    End If
    vbEnv.CommandBars.FindControl(msoControlButton, 2553).Execute
    DoEvents '        vbEnv.CommandBars.FindControl(msoControlButton, 548).Execute
    apiSleep 25
    vbEnv.CommandBars.FindControl(msoControlButton, 642).Execute
End Sub
Private Sub m_apcCommandBarEventProperties_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    If vbEnv.ActiveVBProject Is Nothing Then
        MsgBox "You need an active project to change properties"
        Exit Sub
    End If
    frmProjectProperties.Show
End Sub
Private Sub m_apcCommandBarEventRun_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    If vbEnv.ActiveVBProject Is Nothing Then
        MsgBox "You need an active project to debug"
        Exit Sub
    End If
    If InStr(1, vbEnv.SelectedVBComponent.CodeModule.CodePane.Window.caption, "Module") <> 0 Or InStr(1, vbEnv.SelectedVBComponent.CodeModule.CodePane.Window.caption, "Class") <> 0 Then
        vbEnv.ActiveVBProject.VBComponents.Item(1).Activate
        DoEvents
        apiSleep 25
    End If
    vbEnv.CommandBars.FindControl(msoControlButton, 186).Execute
End Sub
Private Sub m_apcCommandBarEventStart_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    If vbEnv.ActiveVBProject Is Nothing Then
        MsgBox "You need an active project to start a compiled file"
        Exit Sub
    End If
    vbEnv.CommandBars("Standard").Controls("Start With Full Compile").enabled = False
    Dim pth As String
    Dim apd As String
    apd = GetProjectDirectory(vbEnv.ActiveVBProject)
    If Dir(apd, vbNormal) = "" Then MkDir (apd)
    MakeExe True, vbEnv.ActiveVBProject
    DoEvents
    apiSleep 400
    pth = apd & "\" & vbEnv.ActiveVBProject.Name & ".exe"
    If Dir(pth, vbNormal) = "" Then
        vbEnv.CommandBars("Standard").Controls("Start With Full Compile").enabled = True
        Exit Sub
    End If
    Dim ret As Long
    ret = apiWinExec(pth, 1)
    vbEnv.CommandBars("Standard").Controls("Start With Full Compile").enabled = True
    If ret = ERROR_OUT_OF_MEM Then
        MsgBox "Out of memory..."
    ElseIf ERROR_SUCCESS Then
        'do nothing, we're good
    ElseIf ERROR_BAD_FORMAT Then
        MsgBox "Out of memory..."
    ElseIf ERROR_FILE_NOT_FOUND Then
        MsgBox "File not found..."
    ElseIf ERROR_PATH_NOT_FOUND Then
        MsgBox "Path not found..."
    ElseIf ERROR_NO_ASSOC Then
        MsgBox "No file association..."
    Else
        MsgBox "Unknown error.  Perhaps, an access violation without proper rights to write to disk"
    End If
End Sub
Private Sub m_apcCommandBarEventBreak_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    If vbEnv.ActiveVBProject Is Nothing Then
        MsgBox "You need an active project to pause debugging"
        Exit Sub
    End If
    vbEnv.CommandBars.FindControl(msoControlButton, 189).Execute
End Sub
Private Sub m_apcCommandBarEventEnd_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    If vbEnv.ActiveVBProject Is Nothing Then
        MsgBox "You need an active project to stop debugging"
        Exit Sub
    End If
    vbEnv.CommandBars.FindControl(msoControlButton, 228).Execute
End Sub
Private Sub m_apcCommandBarEventAddProcedure_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    If vbEnv.ActiveVBProject Is Nothing Then
        MsgBox "You need an active project to add a procedure"
        Exit Sub
    End If
    If InStr(1, vbEnv.ActiveWindow.caption, "(Code)") = 0 Then
        vbEnv.CommandBars.FindControl(msoControlButton, 2558).Execute
        DoEvents
        apiSleep (25)
    End If
    vbEnv.CommandBars.FindControl(msoControlButton, 559).Execute
End Sub
Private Sub m_apcCommandBarEventOptions_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    If vbEnv.ActiveVBProject Is Nothing Then
        MsgBox "You need an active project change options"
        Exit Sub
    End If
    vbEnv.CommandBars.FindControl(msoControlButton, 522).Execute
End Sub
Private Sub m_apcCommandBarEventToolbox_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    vbEnv.CommandBars.FindControl(msoControlButton, 2553).Execute
    DoEvents
    apiSleep 25
    vbEnv.CommandBars.FindControl(msoControlButton, 548).Execute
End Sub
Private Sub m_apcCommandBarEventBack_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    MsgBox "Feature not implemented yet"
End Sub
Private Sub m_apcCommandBarEventMetric_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    frmCodeMetricAnal.Show
End Sub
Private Sub m_apcCommandBarEventCodeFormat_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    frmCodeFormat.Show
End Sub
Private Sub m_apcCommandBarEventCodeToggleExpansion_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    CollapseExpandCodeProcedure vbEnv.SelectedVBComponent.CodeModule, False
End Sub
Private Sub m_apcCommandBarEventCodeCollapseAll_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    toggleCodeModule = Not toggleCodeModule
    CollapseExpandCodeModule
    Dim vbComBarButton As Object
    Set vbComBarButton = vbEnv.CommandBars("Standard").Controls("Toggle Definitions")
    If toggleCodeModule = True Then
       ' vbComBarButton.Picture = LoadPicture(app.Path & "\images\CollapseDefinitions.bmp")
         vbComBarButton.Picture = frmVB6x.ImageList1.ListImages.Item(5).Picture
    
    Else
       ' vbComBarButton.Picture = LoadPicture(app.Path & "\images\ToggleExpansion.bmp")
           vbComBarButton.Picture = frmVB6x.ImageList1.ListImages.Item(6).Picture
    
    End If
    Set vbComBarButton = Nothing
End Sub
Private Sub CollapseExpandCodeModule(Optional ByVal vbcodmod As VBIDE.CodeModule = Nothing)
    On Error Resume Next
    If vbcodmod Is Nothing Then Set vbcodmod = vbEnv.SelectedVBComponent.CodeModule
    If Not vbcodmod Is Nothing Then
        Dim collapsedlines As String
        Dim procName       As Variant
        Dim procs()        As String
        Dim psl            As Long
        Dim pel            As Long
        Dim v              As Long
        Dim expandedlines  As String
        Dim lns()          As String
        Dim ln             As Variant
        Dim exline         As String
        Dim exp            As String
        Dim sv             As String
        Dim b              As Boolean
        Dim plines         As String
        Dim bl             As Long
        Dim z              As Long
        Dim bCollapsed     As Boolean
        Dim plin           As String
        vbcodmod.CodePane.SetSelection 1, 1, 1, 1
        procs = ListProcedures(vbcodmod)
        For Each procName In procs
            If Trim(procName) <> "" Then
                psl = vbcodmod.ProcStartLine(procName, vbext_pk_Proc)
                pel = psl + vbcodmod.ProcCountLines(procName, vbext_pk_Proc) - 1
                bl = vbcodmod.ProcBodyLine(procName, vbext_pk_Proc)
                vbcodmod.CodePane.SetSelection psl, 1, psl, 1
                For z = bl To pel
                    plin = vbcodmod.Lines(z, 1)
                    plin = ConvertApostropheToChr(plin).lcode
                    plin = ConvertColonToChr(plin)
                    If InStr(1, plin, ": ") <> 0 And IsGoToPoint(plin) = False Then
                        bCollapsed = True
                        Exit For
                    End If
                Next
                If toggleCodeModule = False Then 'expand procedures
                    If bCollapsed = True Then
                        ExpandCodeProcedure psl, pel, vbcodmod
                        FormatCodeModule vbcodmod, True
                    End If
                Else 'collapse procedures
                    If bCollapsed = False Then
                        CollapseCodeProcedure psl, pel, vbcodmod
                        FormatCodeModule vbcodmod, True
                    End If
                End If
            End If
        Next
        Set vbcodmod = Nothing
    End If
End Sub
Friend Sub CollapseExpandCodeProcedure(Optional ByVal vbcodmod As VBIDE.CodeModule = Nothing, Optional doubleclick As Boolean = False)
    On Error Resume Next
    If vbcodmod Is Nothing Then Set vbcodmod = vbEnv.SelectedVBComponent.CodeModule
    If Not vbcodmod Is Nothing Then
        Dim h1 As Long
        h1 = apiLoadCursor(ByVal 0&, IDC_APPSTARTING)
        apiSetSystemCursor h1, OCR_NORMAL
        Dim sl As Long
        Dim sc As Long
        Dim el As Long
        Dim ec As Long
        vbcodmod.CodePane.GetSelection sl, sc, el, ec
        Dim procName As String
        procName = vbcodmod.ProcOfLine(sl, vbext_pk_Proc)
        If Trim(procName) <> "" Then
            Dim psl As Long
            Dim pel As Long
            Dim bl  As Long
            psl = vbcodmod.ProcStartLine(procName, vbext_pk_Proc)
            pel = psl + vbcodmod.ProcCountLines(procName, vbext_pk_Proc) - 1
            bl = vbcodmod.ProcBodyLine(procName, vbext_pk_Proc)
            If (doubleclick = True And sl = bl And sl = el And Mid(vbcodmod.Lines(sl, 1), sc, ec - sc) = procName) Or doubleclick = False Then
                Dim b          As Boolean
                Dim sv         As String
                Dim z          As Long
                Dim bCollapsed As Boolean
                Dim plin       As String
                For z = bl To pel
                    plin = vbcodmod.Lines(z, 1)
                    plin = ConvertApostropheToChr(plin).lcode
                    plin = ConvertColonToChr(plin)
                    If InStr(1, plin, ": ") <> 0 And IsGoToPoint(plin) = False Then
                        bCollapsed = True
                        Exit For
                    End If
                Next
                Dim vbComBarButton As Object
                Set vbComBarButton = vbEnv.CommandBars("Code Window").Controls("Toggle Procedure &Expansion")
                If bCollapsed = True Then
                    ExpandCodeProcedure psl, pel, vbcodmod
                    FormatCodeModule vbcodmod, True
                    vbcodmod.CodePane.SetSelection psl + 1, 1, psl + 1, 1
                    'vbComBarButton.Picture = LoadPicture(app.Path & "\images\ToggleExpansion.bmp")
                    vbComBarButton.Picture = frmVB6x.ImageList1.ListImages.Item(5).Picture
               Else
                    CollapseCodeProcedure psl, pel, vbcodmod
                    FormatCodeModule vbcodmod, True
                    vbcodmod.CodePane.SetSelection psl, 1, psl, 1
                   ' vbComBarButton.Picture = LoadPicture(app.Path & "\images\CollapseDefinitions.bmp")
                     vbComBarButton.Picture = frmVB6x.ImageList1.ListImages.Item(6).Picture
                End If
                Set vbComBarButton = Nothing
            End If
        End If
        Set vbcodmod = Nothing
        apiSetSystemCursor h1, OCR_NORMAL
    End If
End Sub
Friend Function IgnoreError(ByVal er As String, ByVal ptxt As String) As Boolean
    Dim b As Boolean
    If er = "Invalid in Immediate pane" Then er = ""
    If er = "Variable not yet created in this context" Then er = ""
    If er = "Block If without End If" And cfInsertEnd = True Then er = ""
    If er = "End If without block If" And cfInsertEnd = True Then er = ""
    If er = "Else without If" And cfInsertEnd = True Then er = ""
    If er = "Do without Loop" And cfInsertEnd = True Then er = ""
    If er = "Loop without Do" And cfInsertEnd = True Then er = ""
    If er = "For without Next" And cfInsertEnd = True Then er = ""
    If er = "Next without For" And cfInsertEnd = True Then er = ""
    If er = "While without Wend" And cfInsertEnd = True Then er = ""
    If er = "Wend without While" And cfInsertEnd = True Then er = ""
    If er = "Select Case without End Select" And cfInsertEnd = True Then er = ""
    If er = "End Select without Select Case" And cfInsertEnd = True Then er = ""
    If er = "Case without Select Case" And cfInsertEnd = True Then er = ""
    If er = "Expected End With" And cfInsertEnd = True Then er = ""
    If er = "End With without With" And cfInsertEnd = True Then er = ""
    If er = "Sub or Function not defined" Then
        Dim prs() As String
        prs = apc.ListProcedures(vbEnv.SelectedVBComponent.CodeModule)
        If InStr(1, ptxt, " ") <> 0 Then ptxt = Left(ptxt, InStr(1, ptxt, " "))
        If InStr(1, ptxt, "(") <> 0 Then ptxt = Left(ptxt, InStr(1, ptxt, "("))
        If InStr(1, ptxt, Chr(34)) <> 0 Then ptxt = Left(ptxt, InStr(1, ptxt, Chr(34)))
        ptxt = Trim(ptxt)
        Dim prc As Variant
        For Each prc In prs
            prc = Trim(prc)
            If LCase(prc) = LCase(ptxt) Then
                er = ""
                Exit For
            End If
        Next
    End If
    If er = "" Then b = True
    IgnoreError = b
End Function
Private Function CollapseCodeProcedure(ByVal psl As Long, ByVal pel As Long, ByVal vbcodmod As VBIDE.CodeModule) As Boolean
    On Error Resume Next
    Dim expandedline   As String
    Dim collapsedlines As String
    Dim v              As Long
    Dim plines         As String
    plines = Trim(LCase(vbcodmod.Lines(psl, pel)))
    Dim comms        As String
    Dim continuation As String
    Dim continue     As Boolean
    Dim precom       As String
    Dim second       As Long
    Dim cline        As String
    Dim gend         As Boolean
    For v = psl To pel
        cline = Trim(vbcodmod.Lines(v, 1))
        If (Left(cline, 1) = "'" Or cline = "") And gend = False Then
            psl = psl + 1
        Else
            gend = True
            If Left(LCase(cline), 7) = "end sub" Or Left(LCase(cline), 12) = "end function" Then
                pel = v
                Exit For
            End If
        End If
    Next
    Dim continued  As String
    Dim pl         As String
    Dim selcomment As String
    For v = psl To pel
        expandedline = Trim(vbcodmod.Lines(v, 1))
        If InStr(1, UCase(" " & expandedline), " REM ") <> 0 Then
            collapsedlines = ""
            Exit For
        End If
        Dim lc As LineComment
        If Right(Trim(expandedline), 2) = " _" Then
            expandedline = Trim(Left(Trim(expandedline), Len(Trim(expandedline)) - 2))
            continue = True
            continued = continued & "_" & Len(expandedline)
            lc = ConvertApostropheToChr(expandedline)
            expandedline = ConvertColonToChr(lc.lcode)
        Else
            If continue = True Then continued = continued & "_" & Len(expandedline)
            lc = ConvertApostropheToChr(expandedline)
            expandedline = ConvertColonToChr(lc.lcode)
            continue = False
            pl = Trim(LCase(lc.lcode))
            If Left(pl, 12) = "select case " Then
                selcomment = lc.lcomment
                lc.lcomment = ""
            Else
                If selcomment <> "" Then
                    lc.lcomment = selcomment & " - " & lc.lcomment
                    selcomment = ""
                End If
                If Len(pl) >= 5 And Left(pl, 3) = "if " And Right(pl, 5) <> " then" Then
                    pl = ConvertIfToSelectCase(pl)
                    comms = comms & "' ' "
                End If
                If Trim(expandedline) = "" Then
                    If comms <> "" And lc.lcomment <> "" Then
                        comms = comms & " | " & lc.lcomment
                    End If
                Else
                    comms = comms & "' " & continued & " " & lc.lcomment
                End If
            End If
            continued = ""
        End If
        If InStr(1, expandedline, ":") <> 0 And Left(LCase(expandedline), 4) <> "dim " And IsGoToPoint(expandedline) = False Then
            expandedline = Replace(expandedline, ":", Chr(34) & Chr(38) & Chr(34) & Chr(58) & Chr(34) & Chr(38) & Chr(34))
            expandedline = Replace(expandedline, Chr(34) & Chr(34) & Chr(38), "")
            expandedline = Replace(expandedline, Chr(38) & Chr(34) & Chr(34), "")
        End If
        If InStr(1, LCase(expandedline), "case else") = 1 Then expandedline = "Case True"
        Dim er As String
        If apc.ApcHost.Projects(vbEnv.ActiveVBProject.Name).ParseLine(expandedline) = True Then
            er = ""
        Else
            er = apc.ApcHost.Projects(vbEnv.ActiveVBProject.Name).ParseLineError
        End If
        If IgnoreError(er, expandedline) = True Then er = ""
        '        If er = "Invalid in Immediate pane" Then er = ""
        '        If er = "Variable not yet created in this context" Then er = ""
        '        If er = "Block If without End If" And cfInsertEnd = True Then er = ""
        '        If er = "End If without block If" And cfInsertEnd = True Then er = ""
        '        If er = "Else without If" And cfInsertEnd = True Then er = ""
        '        If er = "Do without Loop" And cfInsertEnd = True Then er = ""
        '        If er = "Loop without Do" And cfInsertEnd = True Then er = ""
        '        If er = "For without Next" And cfInsertEnd = True Then er = ""
        '        If er = "Next without For" And cfInsertEnd = True Then er = ""
        '        If er = "While without Wend" And cfInsertEnd = True Then er = ""
        '        If er = "Wend without While" And cfInsertEnd = True Then er = ""
        '        If er = "Select Case without End Select" And cfInsertEnd = True Then er = ""
        '        If er = "End Select without Select Case" And cfInsertEnd = True Then er = ""
        '        If er = "Case without Select Case" And cfInsertEnd = True Then er = ""
        '        If er = "Expected End With" And cfInsertEnd = True Then er = ""
        '        If er = "End With without With" And cfInsertEnd = True Then er = ""
        If er <> "" Then
            Dim j As Long
            For j = 1 To vbEnv.CommandBars("Menu Bar").Controls.Count
                If Left(vbEnv.CommandBars("Menu Bar").Controls.Item(j).caption, 13) = "Syntax Check:" Then
                    vbEnv.CommandBars("Menu Bar").Controls.Item(j).caption = "Syntax Check:   " & er
                    Exit For
                End If
            Next
            Exit Function
        End If
        '
        If IsGoToPoint(expandedline) = True Or IsOnErrorGoTo(expandedline) = True Then
            If Len(collapsedlines) > 0 And Right(collapsedlines, 1) = ":" Then collapsedlines = Left(collapsedlines, Len(collapsedlines) - 1)
            If continue = False Then
                collapsedlines = collapsedlines & vbCrLf & Trim(expandedline) & vbCrLf
            Else
                collapsedlines = collapsedlines & Trim(expandedline)
            End If
        Else
            expandedline = ConvertIfToSelectCase(expandedline)
            If expandedline <> "" Then
                collapsedlines = collapsedlines & Trim(expandedline) '& ":"
                If continue = False Then collapsedlines = collapsedlines & ":"
            End If
        End If
    Next
    If Len(collapsedlines) > 0 And Right(collapsedlines, 1) = ":" Then collapsedlines = Left(collapsedlines, Len(collapsedlines) - 1)
    Dim plin    As Variant
    Dim plins() As String
    plins = Split(plines, vbCrLf)
    Dim tmp As String
    For Each plin In plins
        tmp = tmp & Trim(plin) & vbCrLf
    Next
    If collapsedlines <> "" And Replace(LCase(collapsedlines), vbCrLf, "") <> Replace(tmp, vbCrLf, "") Then
        vbcodmod.DeleteLines psl, (pel - psl) + 1
        vbcodmod.InsertLines psl, comms & vbCrLf & collapsedlines
        CollapseCodeProcedure = True
    End If
    Set vbcodmod = Nothing
End Function
Private Function ExpandCodeProcedure(ByVal psl As Long, ByVal pel As Long, ByVal vbcodmod As VBIDE.CodeModule) As Boolean
    On Error Resume Next
    Dim v             As Long
    Dim expandedlines As String
    Dim lns()         As String
    Dim ln            As Variant
    Dim sv            As String
    Dim b             As Boolean
    Dim exline        As String
    Dim u             As Long
    u = 0
    Dim cline         As String
    Dim comms()       As String
    Dim k             As Long
    Dim commentheader As Boolean
    For v = psl To pel
        cline = Trim(vbcodmod.Lines(v, 1))
        If Left(cline, 1) = "'" Or cline = "" Then
            psl = psl + 1
            commentheader = True
        Else
            comms = Split(Trim(vbcodmod.Lines(v - 1, 1)), "'")
            If commentheader = True Then vbcodmod.ReplaceLine v - 1, ""
            Exit For
        End If
    Next
    Dim conts()       As String
    Dim cont          As Variant
    Dim ln2           As String
    Dim topcount      As Long
    Dim colonwcomment As Boolean
    Dim iffy          As String
    Dim enddy         As String
    Dim duddy         As Boolean
    For v = psl To pel
        exline = Trim(vbcodmod.Lines(v, 1))
        If exline <> "" Then
            exline = Replace(exline, Chr(34) & " & Chr(39) & " & Chr(34), "'")
            exline = ConvertColonToChr(exline)
            exline = Trim(exline)
            If IsGoToPoint(exline) = True Then
                exline = Replace(exline, "   ", " ")
                exline = Replace(exline, "  ", " ")
                u = u + 1
                If Trim(comms(u)) <> "" Then comms(u) = "'" & Trim(comms(u))
                If InStr(1, comms(u), " | ") <> 0 Then comms(u) = Replace(comms(u), " | ", vbCrLf & "'")
                expandedlines = expandedlines & exline & comms(u) & vbCrLf
            Else
                If Len(Trim(LCase(exline))) >= 5 And Left(Trim(LCase(exline)), 3) = "if " And InStr(1, Trim(LCase(exline)), " then ") <> 0 Then ' InStr(1, reducespaces(Trim(LCase(exline))), " then Chr(39)") = 0 Then
                    iffy = Left(exline, InStr(1, exline, " Then ") + 4)
                    enddy = Right((exline), Len(exline) - (InStr(1, exline, " Then ") + 5))
                    duddy = True
                    expandedlines = expandedlines & iffy & vbCrLf & enddy & vbCrLf & "End If"
                Else
                    lns = Split(exline, ": ")
                    sv = ""
                    b = True
                    For Each ln In lns
                        ln = Trim(ln)
                        ln = Replace(ln, Chr(34) & Chr(32) & Chr(38) & Chr(32) & Chr(34) & Chr(58) & Chr(34) & Chr(32) & Chr(38) & Chr(32) & Chr(34), ":")
                        ln = ConvertSelectCaseToIfEndIf(ln, sv, b)
                        If Trim(ln) <> "" Then
                            u = u + 1
                            ln = Replace(ln, Chr(34) & " & Chr(58) & " & Chr(34), ":")
                            conts = Split(comms(u), "_")
                            If InStr(1, comms(u), "_") <> 0 Then
                                comms(u) = Right(comms(u), Len(comms(u)) - InStrRev(comms(u), "_"))
                                comms(u) = Right(comms(u), Len(comms(u)) - InStr(comms(u), " "))
                            End If
                            comms(u) = Replace(comms(u), "Chr(39)", "'")
                            comms(u) = Replace(comms(u), "Chr(95)", "_")
                            comms(u) = Replace(comms(u), "Chr(58)", ":")
                            If InStr(1, comms(u), " | ") <> 0 Then comms(u) = Replace(comms(u), " | ", vbCrLf & "'")
                            If Trim(comms(u)) = "'" Then comms(u) = ""
                            If Trim(comms(u)) <> "" Then
                                If InStr(1, comms(u), vbCrLf) <> 0 Then
                                    If Left(Trim(comms(u)), Len(vbCrLf)) = vbCrLf Then
                                        If duddy = True Then
                                            duddy = False
                                            ln = vbCrLf & ln
                                        Else
                                            ln = ln & comms(u)
                                        End If
                                    Else
                                        ln = ln & " '" & comms(u)
                                    End If
                                Else
                                    ln = ln & " '" & comms(u)
                                End If
                            End If
                            ln = Replace(ln, "Chr(58)", ":")
                            ln = Replace(ln, "Chr(39)", "'")
                            ln = Replace(ln, "Chr(95)", "_")
                            ln = Replace(ln, "Chr(124)", "|")
                            If UBound(conts) = 0 Then ReDim conts(0)
                            If UBound(conts) > 0 Then
                                ln2 = ln
                                For Each cont In conts
                                    topcount = topcount + 1
                                    cont = Trim(cont)
                                    Dim comm As String
                                    If InStr(1, cont, " ") > 0 Then
                                        comm = Right(cont, Len(cont) - InStr(1, cont, " "))
                                        cont = Left(cont, InStr(1, cont, " "))
                                    End If
                                    If IsNumeric(cont) = True Then
                                        If topcount = UBound(conts) + 1 Then
                                            expandedlines = expandedlines & Left(ln2, cont + 1) & vbCrLf
                                        Else
                                            expandedlines = expandedlines & Left(ln2, cont + 1) & " _" & vbCrLf
                                        End If
                                        ln2 = Right(ln2, (Len(ln2) - cont) - 1)
                                    End If
                                Next
                            Else
                                expandedlines = expandedlines & ln & vbCrLf
                            End If
                        End If
                    Next
                End If
                '                Dim c As LineComment
                '                c = ConvertApostropheToChr(exline)
                '                If Trim(c.lcomment) <> "" Then exline = Trim(c.lcode) & " '" & c.lcomment
            End If
        End If
    Next
    expandedlines = Left(expandedlines, Len(expandedlines) - Len(vbCrLf))
    If expandedlines <> "" Then
        vbcodmod.DeleteLines psl, (pel - psl) + 1
        vbcodmod.InsertLines psl, expandedlines
        If Trim(vbcodmod.Lines(psl - 1, 1)) = "" Then vbcodmod.DeleteLines psl - 1, 1
        ExpandCodeProcedure = True
    End If
    Set vbcodmod = Nothing
End Function
Private Function ConvertApostropheToChr(ByVal inputline As String) As LineComment
    Dim i          As Long
    Dim inp        As String
    Dim outputline As String
    Dim b          As Boolean
    For i = 1 To Len(inputline)
        inp = Mid(inputline, i, 1)
        If inp = Chr(34) Then
            b = Not b
        ElseIf inp = "'" Then
            If b = True Then
                inp = Chr(34) & "&Chr(39)&" & Chr(34)
            Else
                ConvertApostropheToChr.lcode = outputline
                ConvertApostropheToChr.lcomment = Replace(Trim(Right(inputline, Len(inputline) - i)), "'", "Chr(39)")
                ConvertApostropheToChr.lcomment = Replace(ConvertApostropheToChr.lcomment, "_", "Chr(95)")
                ConvertApostropheToChr.lcomment = Replace(ConvertApostropheToChr.lcomment, ":", "Chr(58)")
                ConvertApostropheToChr.lcomment = Replace(ConvertApostropheToChr.lcomment, "|", "Chr(124)")
                Exit Function
            End If
        End If
        outputline = outputline & inp
    Next
    ConvertApostropheToChr.lcode = outputline
    ConvertApostropheToChr.lcomment = ""
End Function
Private Function ConvertColonToChr(ByVal inputline As String) As String
    Dim i          As Long
    Dim inp        As String
    Dim outputline As String
    Dim b          As Boolean
    For i = 1 To Len(inputline)
        inp = Mid(inputline, i, 1)
        If inp = Chr(34) Then
            b = Not b
        ElseIf inp = ":" Then
            If b = True Then inp = Chr(34) & " & Chr(58) & " & Chr(34)
        End If
        outputline = outputline & inp
    Next
    ConvertColonToChr = outputline
End Function
Private Function ConvertIfToSelectCase(ByVal inputline As String) As String
    Dim exp As String
    exp = Trim(LCase(inputline))
    If exp <> "" Then
        If Left(exp, 3) = "if " Then
            If Right(exp, 5) = " then" Then ' if this is going to have closing "If/End If"
                exp = Trim(Mid(exp, 3, (Len(exp) - 3) - 4))
                inputline = "Select Case True" & ": " & "Case " & exp
            Else 'otherwise it must be inline "If/Then"
                Dim ex As String
                ex = Trim(Mid(exp, 3, (InStr(1, exp, " then ") - 3))) 'TODO update this to make sure this isn't inside a string of the expression
                Dim xcode As String
                xcode = Trim(Right(exp, Len(exp) - (InStr(1, exp, " then ") + 5)))
                inputline = "Select Case True" & ": " & "Case " & ex & ": " & xcode & ": " & "End Select"
            End If
        ElseIf exp = "end if" Or Left(exp, 7) = "end if " Then
            inputline = "End Select"
        ElseIf Left(exp, 7) = "elseif " Then
            exp = Trim(Mid(exp, 7, (Len(exp) - 7) - 4))
            inputline = "Case " & exp
        ElseIf exp = "else" Or Left(exp, 5) = "else " Then
            inputline = "Case True"
        End If
    End If
    ConvertIfToSelectCase = inputline
End Function
Private Function ConvertSelectCaseToIfEndIf(ByVal inputline As String, ByRef vs As String, ByRef b As Boolean) As String
    Dim exp As String
    exp = Trim(LCase(inputline))
    If exp <> "" Then
        If Left(exp, 16) = "select case true" Then
            inputline = ""
            b = False
            vs = ""
        ElseIf Left(exp, 12) = "select case " Then
            b = False
            inputline = ""
            vs = Trim(Right(exp, Len(exp) - 11))
        ElseIf Left(exp, 5) = "case " Then
            If vs <> "" Then
                inputline = "If " & vs & "=" & Right(exp, Len(exp) - 5) & " Then"
            Else
                inputline = "If " & Right(exp, Len(exp) - 5) & " Then"
            End If
            If b = True Then inputline = "Else" & inputline
            b = True
        ElseIf Left(exp, 10) = "end select" Then
            inputline = "End If"
            b = True
            vs = ""
        End If
    End If
    ConvertSelectCaseToIfEndIf = inputline
End Function
Private Function IsGoToPoint(ByVal inputline As String) As Boolean
    Dim i As Long
    i = InStr(1, inputline, ":")
    If i = 0 Then Exit Function
    inputline = Trim(LCase(inputline))
    Dim s As Long
    s = InStr(1, inputline, " ")
    If s = 0 Then IsGoToPoint = True
    inputline = RemoveExtraSpaces(inputline)
    Dim a As Long
    a = InStr(1, inputline, "'")
    If a = s + 1 Then
        IsGoToPoint = True
    End If
End Function
Private Function RemoveExtraSpaces(ByVal pline As String) As String
    Do
        If InStr(1, pline, "  ") < 1 Then Exit Do
        pline = Replace(pline, "  ", " ")
    Loop
    RemoveExtraSpaces = pline
End Function
Private Function IsOnErrorGoTo(ByVal inputline As String) As Boolean
    inputline = Trim(LCase(inputline))
    If Left(inputline, 14) = "on error goto " Then IsOnErrorGoTo = True
End Function
Private Sub m_apcCommandBarEventObjectBrowser_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    vbEnv.CommandBars.FindControl(msoControlButton, 473).Execute
End Sub
Private Sub m_apcCommandBarEventFind_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    vbEnv.CommandBars.FindControl(msoControlButton, 2558).Execute
    DoEvents
    apiSleep 25
    vbEnv.CommandBars.FindControl(msoControlButton, 141).Execute
End Sub
Private Sub m_apcCommandBarEventUndo_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    vbEnv.CommandBars.FindControl(msoControlButton, 128).Execute
    DoEvents
    apiSleep 25
    vbEnv.CommandBars("Standard").Controls("Undo").enabled = vbEnv.CommandBars.FindControl(msoControlButton, 128).enabled
    vbEnv.CommandBars("Standard").Controls("Redo").enabled = vbEnv.CommandBars.FindControl(msoControlButton, 129).enabled
End Sub
Private Sub m_apcCommandBarEventRedo_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    vbEnv.CommandBars.FindControl(msoControlButton, 129).Execute
    DoEvents
    apiSleep 25
    vbEnv.CommandBars("Standard").Controls("Undo").enabled = vbEnv.CommandBars.FindControl(msoControlButton, 128).enabled
    vbEnv.CommandBars("Standard").Controls("Redo").enabled = vbEnv.CommandBars.FindControl(msoControlButton, 129).enabled
End Sub
Private Sub m_apcCommandBarEventComment_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    vbEnv.CommandBars.FindControl(msoControlButton, 192).Execute
End Sub
Private Sub m_apcCommandBarEventUnComment_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    vbEnv.CommandBars.FindControl(msoControlButton, 2552).Execute
End Sub
Private Sub m_apcCommandBarEventFormat_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    Dim h1 As Long
    h1 = apiLoadCursor(ByVal 0&, IDC_APPSTARTING)
    apiSetSystemCursor h1, OCR_NORMAL
    DoEvents
    apiSleep 25
    Dim vbcomp As VBIDE.VBComponent
    For Each vbcomp In vbEnv.ActiveVBProject.VBComponents
        If Not vbcomp Is Nothing Then
            FormatCodeModule vbcomp.CodeModule
            DoEvents
            apiSleep 0
        End If
    Next
    Set vbcomp = Nothing
    apiSetSystemCursor h1, OCR_NORMAL
End Sub
Private Sub m_apcCommandBarEventRun2_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    m_apcCommandBarEventRun_Click CommandBarControl, handled, CancelDefault
End Sub
Private Sub m_apcCommandBarEventBreak2_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    m_apcCommandBarEventBreak_Click CommandBarControl, handled, CancelDefault
End Sub
Private Sub m_apcCommandBarEventEnd2_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    m_apcCommandBarEventEnd_Click CommandBarControl, handled, CancelDefault
End Sub
Private Sub m_apcCommandBarEventSave2_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    m_apcCommandBarEventSave_Click CommandBarControl, handled, CancelDefault
End Sub
Private Sub m_apcCommandBarEventMake2_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    m_apcCommandBarEventMake_Click CommandBarControl, handled, CancelDefault
End Sub
Private Sub m_apcCommandBarEventStart2_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    m_apcCommandBarEventStart_Click CommandBarControl, handled, CancelDefault
End Sub
'
'Object model events
'HOST
Private Sub m_apcHost_AfterPause()
    'MsgBox "after pause"
End Sub
Private Sub m_apcHost_Activate(Activate As Boolean)
    ' MsgBox "host activated"
End Sub
Private Sub m_apcHost_AfterModalDialog(ByVal Project As MSAPC.IApcProject)
    ' Project.ExecuteLine ("Msgbox" & Chr(34) & "hello" & Chr(34))
End Sub
Private Sub m_apcHost_AfterRun()
    'MsgBox "after run"
End Sub
Private Sub m_apcHost_BeforeModalDialog(ByVal Project As MSAPC.IApcProject)
    ' MsgBox "before modal dialog " & Project.Name
End Sub
Private Sub m_apcHost_BeforePause()
    'MsgBox "before pause"
End Sub
Private Sub m_apcHost_BeforeRun()
    ' MsgBox "before run"
End Sub
Private Sub m_apcHost_CreateMacro(ByVal ProcedureName As String, ByVal CurrentScope As MSAPC.IApcScope, ByVal Description As String, Created As Boolean)
    'MsgBox ProcedureName & " " & Description
End Sub
Private Sub m_apcHost_FindFile(ByVal FileName As String, fullpath As String)
    'MsgBox FileName & " " & FullPath
End Sub
Private Sub m_apcHost_HostCheckReference(ByVal Saving As Boolean, ByVal LibraryGuid As String, MajorVersion As Long, MinorVersion As Long)
    ' MsgBox Saving & " " & LibraryGuid & " " & MajorVersion & " " & MinorVersion
End Sub
Private Sub m_apcHost_NameChange(ByVal Project As MSAPC.IApcProject, ByVal OldName As String)
    On Error Resume Next
    '  MsgBox "Host name changed " & Project.Name
    savingall = True
    Dim i As Long
    For i = 1 To UBound(pSettings)
        If pSettings(i).Name = OldName Then 'if project name has changed
            pSettings(i).Name = Project.Name 'then update the old settings file with the new name
            Exit For
        End If
    Next
End Sub
Private Sub m_apcHost_OpenProject(ByVal ProjectFileName As String, ByVal ReferencingProject As MSAPC.IApcProject, Project As MSAPC.IApcProject)
    'MsgBox "Open project " & ProjectFileName
End Sub
Private Sub m_apcHost_ShowHelp(ByVal HelpFileName As String, ByVal Command As Long, ByVal CommandData As Long, ByVal IsWinHelp As Boolean, pbShowHelp As Boolean)
    ' MsgBox "Show help " & HelpFileName
End Sub
'
'PROJECT
Private Sub m_ApcProject_Save()
    SaveProjects 'Save all open projects
End Sub
Private Sub m_ApcProject_ModuleChange(ByVal ProjectItem As MSAPC.IApcProjectItem)
    '  MsgBox "Module change: " & ProjectItem.Name
    savingall = True
End Sub
Private Sub m_ApcProject_ModuleDirtyChange(ByVal ProjectItem As MSAPC.IApcProjectItem, ByVal Dirty As Boolean)
    isDirty = True
End Sub
Private Sub m_ApcProject_ProjectItemDelete(ByVal ProjectItem As MSAPC.IApcProjectItem)
    '  MsgBox "Project item delete: " & ProjectItem.Name
End Sub
Private Sub m_ApcProject_Reset()
    ' MsgBox "stop/reset"
End Sub
Private Sub m_ApcProject_ModeChange(ByVal NewMode As MSAPC.AxProjectMode)
    '  MsgBox "project mode changed " & NewMode
End Sub
Private Sub m_ApcProject_NameChange(ByVal ProjectItem As MSAPC.IApcProjectItem, ByVal OldName As String)
    'This event fires when importing a file, which by default is given a cardinal index with a generic prefix
    ' MsgBox "name changed " & OldName & " to " & ProjectItem.Name
    savingall = True
End Sub
Private Sub m_ApcProject_OpenProjectItem(ByVal Name As String, ByVal id As Long, HostProjectItem As Object)
    'MsgBox "opened projet item " & Name
    savingall = True
End Sub
Private Sub m_ApcProject_ProcedureAdd(ByVal Procedure As MSAPC.IApcProcedure)
    ' MsgBox "procedure added " & Procedure.Description
End Sub
Private Sub m_ApcProject_ProcedureDelete(ByVal ProjectItem As MSAPC.IApcProjectItem, ByVal ProcedureName As String, ByVal ProcedureType As MSAPC.AxProcedureType)
    ' MsgBox "Procedure delete " & ProjectItem.Name
End Sub
Private Sub m_ApcProject_ProjectItemCreated(ByVal ProjectItem As MSAPC.IApcProjectItem)
    On Error Resume Next
    savingall = True
    If vbEnv.ActiveVBProject Is Nothing Then Exit Sub
    If ProjectItem.Type = axTypeDesigner Then
        Dim i As Long
        Dim z As Long
        For i = 1 To UBound(pSettings)
            If vbEnv.ActiveVBProject.Name = pSettings(i).Name Then
                z = i
                Exit For
            End If
        Next
        If Replace(pSettings(z).Startup, Chr(34), "") = "" Then pSettings(z).Startup = Chr(34) & ProjectItem.Name & Chr(34)
        If Replace(pSettings(z).IconForm, Chr(34), "") = "" Then pSettings(z).IconForm = Chr(34) & ProjectItem.Name & Chr(34)
    End If
End Sub
Private Sub m_apcReferenceEvents_ItemAdded(ByVal Reference As VBIDE.Reference)
    MsgBox ReferenceChange("Reference Added " & Reference.Name, Reference.Type)
End Sub
Private Sub m_apcReferenceEvents_ItemRemoved(ByVal Reference As VBIDE.Reference)
    MsgBox ReferenceChange("Reference Removed " & Reference.Name, Reference.Type)
End Sub
Private Function ReferenceChange(ByVal added As String, ByVal t As Long) As String
    If t = vbext_rk_TypeLib Then
        ReferenceChange = "Type Library "
    ElseIf t = vbext_rk_Project Then
        ReferenceChange = "Project "
    End If
    ReferenceChange = ReferenceChange & added
End Function
Private Sub InitializeProjectSettings(ByRef p As PROJECTSETTINGS) 'Mirror of the project file .vbp
    On Error Resume Next
    With p
        .ProjectType = "exe"
        ReDim Preserve .References(0)
        ReDim Preserve .Objects(0)
        ReDim Preserve .Modules(0)
        ReDim Preserve .UserControls(0)
        ReDim Preserve .Forms(0)
        ReDim Preserve .Classes(0)
        ReDim Preserve .RelatedDoc(0)
        ReDim Preserve .ResFile32(0)
        ReDim Preserve .Unknown(0)
        .IconForm = Chr(34) & Chr(34)
        .Startup = Chr(34) & "" & Chr(34)
        .HelpFile = Chr(34) & Chr(34)
        .title = Chr(34) & Chr(34)
        .ExeName32 = Chr(34) & Chr(34)
        .Path32 = Chr(34) & Chr(34)
        .Command32 = Chr(34) & Chr(34)
        .Name = Chr(34) & Chr(34)
        .HelpContextID = Chr(34) & "0" & Chr(34)
        .Description = Chr(34) & Chr(34)
        .CompatibleMode = Chr(34) & "0" & Chr(34)
        .MajorVer = "1"
        .MinorVer = "0"
        .RevisionVer = "0"
        .AutoIncrementVer = "0"
        .ServerSupportFiles = "0"
        .VersionComments = Chr(34) & Chr(34)
        .VersionCompanyName = Chr(34) & Chr(34)
        .VersionFileDescription = Chr(34) & Chr(34)
        .VersionLegalCopyright = Chr(34) & Chr(34)
        .VersionLegalTrademarks = Chr(34) & Chr(34)
        .VersionProductName = Chr(34) & Chr(34)
        .CompilationType = "0"
        .OptimizationType = "0"
        .FavorPentiumPro = "0"
        .CodeViewDebugInfo = "0"
        .NoAliasing = "0"
        .BoundsCheck = "0"
        .OverflowCheck = "0"
        .FlPointCheck = "0"
        .FDIVCheck = "0"
        .UnroundedFP = "0"
        .StartMode = "0"
        .Unattended = "0"
        .Retained = "0"
        .ThreadPerObject = "0"
        .MaxNumberOfThreads = "1"
        .DebugStartupOption = "0"
        '.AutoRefresh = "0"
    End With
End Sub
Friend Function CreateNewProject(Optional ByVal pname As String = "") As VBProject
    On Error Resume Next
    savingall = True
    Dim cpn As String
    Dim b   As Boolean
    If pname = "" Then
        cpn = RandomProjectName
        b = True
    Else
        cpn = pname
    End If
    Dim vb As VBProject
    For Each vb In vbEnv.VBProjects
        If vb.Name = cpn Then
            Exit Function
        End If
    Next
    Set m_ApcProject = m_apcHost.Projects.Add(axProjectDisableSave Or axProjectThrowAwayCompiledState, cpn)   ' Create a default projectaxProjectNormal  'axProjectThrowAwayCompiledState
    m_ApcProject.Dirty = False ' Set the proj5ect's dirty flag to False so no save prompt will occur unless user modifications are made/projects created etc
    ReDim Preserve pSettings(UBound(pSettings) + 1)
    InitializeProjectSettings pSettings(UBound(pSettings))
    pSettings(UBound(pSettings)).Name = cpn
    pSettings(UBound(pSettings)).ExeName32 = Chr(34) & cpn & ".exe" & Chr(34)
    If b = True And Not vbEnv.ActiveVBProject Is Nothing And vbEnv.VBProjects.Count = 1 Then
        vbEnv.CommandBars.FindControl(msoControlButton, 512).Execute
    End If
    Set CreateNewProject = apc.ApcProject.VBProject
End Function
Friend Sub OpenLoadProject()
    On Error Resume Next
    Dim fso  As Object
    Dim f    As String
    Dim filt As String
    filt = "Visual Basic Project Files (*.vbp; *.vbg)" & Chr(0) & "*.vbp;*.vbg" & Chr(0)
    f = FileOpenSave(0, filt, 1, "vbp", "", "Open Visual Basic Project File", -1, True)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.fileexists(f) = False Then
        Set fso = Nothing
        Exit Sub
    End If
    Dim ext As String
    ext = fso.GetExtensionName(f)
    If LCase(ext) = "vbp" Then
        Call LoadProject(f)
    ElseIf LCase(ext) = "vbg" Then
        Call OpenProjectGroup(f)
    End If
    Set fso = Nothing
End Sub
Friend Sub OpenProjectGroup(ByVal f As String)
    On Error Resume Next
    lastgroup = f
    Dim fso     As Object
    Dim otf     As Object
    Dim inp     As String
    Dim projs() As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set otf = fso.OpenTextFile(f, 1, 0)
    Dim b As Boolean
    Do
        If otf.AtEndOfStream = True Then Exit Do
        inp = otf.Readline
        projs = Split(inp, "=")
        b = False
        Dim vb As VBProject
        For Each vb In vbEnv.VBProjects
            If vb.Name = Replace(projs(1), ".vbp", "") Or vb.Name = Replace(projs(1), ".vbg", "") Then
                b = True
                Exit For
            End If
        Next
        If b = False And UBound(projs) = 1 Then
            LoadProject app.Path & "\Projects\" & projs(1)
        End If
    Loop
    otf.Close
    Set fso = Nothing
End Sub
Friend Function FileOpenSave(ByRef flags As Long, Optional ByVal Filter As String = vbNullString, Optional ByVal FilterIndex As Long = 1, Optional ByVal DefaultExt As String = vbNullString, Optional ByVal FileName As String = vbNullString, Optional ByVal DialogTitle As String = vbNullString, Optional ByVal hwnd As Long = -1, Optional ByVal OpenFile As Boolean = True) As String
    On Error Resume Next
    Dim ofn         As OPENFILENAME
    Dim stFileName  As String
    Dim stFileTitle As String
    Dim fResult     As Long
    If (hwnd = -1) Then hwnd = 0
    stFileName = Left(FileName & String(260, vbNullChar), 260)
    stFileTitle = String(260, vbNullChar)
    With ofn
        .lStructSize = Len(ofn)
        .hwndOwner = hwnd
        .stFilter = Filter
        .nFilterIndex = FilterIndex
        .strFile = stFileName
        .nMaxFile = Len(stFileName)
        .stFileTitle = stFileTitle
        .nMaxFileTitle = Len(stFileTitle)
        .strTitle = DialogTitle
        .flags = flags
        .stDefExt = DefaultExt
        .stInitialDir = app.Path & "\Projects"
        .hInstance = 0
        .stCustomFilter = String(260, vbNullChar)
        .nMaxCustFilter = 260
        .lpfnHook = 0
    End With
    If OpenFile = True Then
        fResult = apiGetOpenFileName(ofn)
    Else
        fResult = apiGetSaveFileName(ofn)
    End If
    Dim of As OPENFILENAME
    Call apiGetOpenFileName(of) 'clear initial directory every time otherwise it will wander around
    If fResult = 0 Then Exit Function
    flags = ofn.flags
    FileOpenSave = Left(ofn.strFile, InStr(1, ofn.strFile, vbNullChar, vbBinaryCompare) - 1)
End Function
Friend Sub LoadProject(ByVal pth As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.fileexists(pth) = False Then Set fso = Nothing: Exit Sub
    Dim sr     As Object
    Dim pname  As String
    Dim th     As String
    Dim vbproj As New VBIDE.VBProject
    Dim vbcomp As New VBIDE.VBComponent
    pname = fso.GetBaseName(pth)
    CreateNewProject pname
    Dim i As Long
    Dim z As Long
    z = 1
    For i = 1 To UBound(pSettings)
        If pSettings(i).Name = pname Then
            z = i
            Exit For
        End If
    Next
    Dim hashes() As String
    Dim gui      As String
    Dim maj      As String
    Dim min      As String
    Dim vers()   As String
    Dim fslash() As String
    Dim b        As Boolean
    Dim n As Long
    
    Set sr = fso.OpenTextFile(pth, 1, False, 0)
    Do 'TO-DO frm ctl pag dsr bas cls res
        If sr.AtEndOfStream = True Then Exit Do
        Dim inp As String
        inp = Trim(sr.Readline)
        If InStr(1, inp, "=") Then
            Dim pv() As String
            pv = Split(inp, "=")
            If pv(0) = "Type" Then
                pSettings(z).ProjectType = pv(1)
            ElseIf pv(0) = "Reference" Then
                ReDim Preserve pSettings(z).References(UBound(pSettings(z).References) + 1)
                pSettings(z).References(UBound(pSettings(z).References)) = pv(1)
                hashes = Split(pv(1), "#")
                If UBound(hashes) > 2 Then vbEnv.ActiveVBProject.References.AddFromFile fso.getfilename(hashes(3))
            ElseIf pv(0) = "Object" Then
                ReDim Preserve pSettings(z).Objects(UBound(pSettings(z).Objects) + 1)
                pSettings(z).Objects(UBound(pSettings(z).Objects)) = pv(1)
                If InStr(1, pv(1), "{") <> 0 Then pv(1) = Right(pv(1), (1 + Len(pv(1))) - InStr(1, pv(1), "{"))
                hashes = Split(pv(1), "#")
                If UBound(hashes) > 1 Then fslash = Split(hashes(2), ";")
                If UBound(fslash) > 0 Then
                    b = False
                    For n = 1 To vbEnv.ActiveVBProject.References.Count
                        If InStr(1, vbEnv.ActiveVBProject.References.Item(n).fullpath, Trim(fslash(1))) <> 0 Then
                            b = True
                            Exit For
                        End If
                    Next
                    If b = False Then vbEnv.ActiveVBProject.References.AddFromFile fso.getfilename(Trim(fslash(1)))
                End If
                '                gui = ""
                '                maj = ""
                '                min = ""
                '                If UBound(hashes) > 0 Then
                '                    gui = hashes(0)
                '                    If InStr(1, gui, "{") <> 0 Then gui = Right(gui, (1 + Len(gui)) - InStr(1, gui, "{"))
                '                    vers = Split(hashes(1), ".")
                '                    If UBound(vers) > 0 Then
                '                        maj = vers(0)
                '                        min = vers(1)
                '                       ' MsgBox gui & vbCrLf & maj & vbCrLf & min
                '                       vbEnv.ActiveVBProject.References.AddFromGuid gui, maj, min
                '                    End If
                '                End If
            ElseIf pv(0) = "Module" Then
                If InStr(1, pv(1), ";") <> 0 Then
                    th = Right(pv(1), Len(pv(1)) - InStr(1, pv(1), ";") - 1)
                    If fso.fileexists(th) = True Then
                        Set vbproj = vbEnv.ActiveVBProject
                        Set vbcomp = vbproj.VBComponents.Import(th)
                        Set vbproj = Nothing
                        Set vbcomp = Nothing
                    End If
                End If
            ElseIf pv(0) = "UserControl" Then
                ReDim Preserve pSettings(z).UserControls(UBound(pSettings(z).UserControls) + 1)
                pSettings(z).UserControls(UBound(pSettings(z).UserControls)) = pv(1)
            ElseIf pv(0) = "Form" Then
            ElseIf pv(0) = "Designer" Then
                th = pv(1)
                If fso.fileexists(th) = True Then
                    Set vbproj = vbEnv.ActiveVBProject
                    Set vbcomp = vbproj.VBComponents.Import(th)
                    Set vbproj = Nothing
                    Set vbcomp = Nothing
                End If
            ElseIf pv(0) = "Class" Then
                If InStr(1, pv(1), ";") <> 0 Then
                    th = Right(pv(1), Len(pv(1)) - InStr(1, pv(1), ";") - 1)
                    If fso.fileexists(th) = True Then
                        Set vbproj = vbEnv.ActiveVBProject
                        Set vbcomp = vbproj.VBComponents.Import(th)
                        Set vbproj = Nothing
                        Set vbcomp = Nothing
                    End If
                End If
            ElseIf pv(0) = "ResFile32" Then
                ReDim Preserve pSettings(z).ResFile32(UBound(pSettings(z).ResFile32) + 1)
                pSettings(z).ResFile32(UBound(pSettings(z).ResFile32)) = pv(1)
            ElseIf pv(0) = "RelatedDoc" Then
                ReDim Preserve pSettings(z).RelatedDoc(UBound(pSettings(z).RelatedDoc) + 1)
                pSettings(z).RelatedDoc(UBound(pSettings(z).RelatedDoc)) = pv(1)
            ElseIf pv(0) = "IconForm" Then
                pSettings(z).IconForm = pv(1)
            ElseIf pv(0) = "Startup" Then
                pSettings(z).Startup = pv(1)
            ElseIf pv(0) = "HelpFile" Then
                pSettings(z).HelpFile = Replace(pv(1), Chr(34), "")
                vbEnv.VBProjects.Item(z).HelpFile = pSettings(z).HelpFile
            ElseIf pv(0) = "Title" Then
                pSettings(z).title = pv(1)
            ElseIf pv(0) = "ExeName32" Then
                pSettings(z).ExeName32 = pv(1)
            ElseIf pv(0) = "Path32" Then
                pSettings(z).Path32 = pv(1)
            ElseIf pv(0) = "Command32" Then
                pSettings(z).Command32 = pv(1)
            ElseIf pv(0) = "Name" Then
                pSettings(z).Name = Replace(pv(1), Chr(34), "")
                vbEnv.VBProjects.Item(z).Name = pSettings(z).Name
            ElseIf pv(0) = "HelpContextID" Then
                pSettings(z).HelpContextID = Replace(pv(1), Chr(34), "")
                vbEnv.VBProjects.Item(z).HelpContextID = pSettings(z).HelpContextID
            ElseIf pv(0) = "Description" Then
                pSettings(z).Description = Replace(pv(1), Chr(34), "")
                vbEnv.VBProjects.Item(z).Description = pSettings(z).Description
            ElseIf pv(0) = "CompatibleMode" Then
                pSettings(z).CompatibleMode = pv(1)
            ElseIf pv(0) = "MajorVer" Then
                pSettings(z).MajorVer = pv(1)
            ElseIf pv(0) = "MinorVer" Then
                pSettings(z).MinorVer = pv(1)
            ElseIf pv(0) = "RevisionVer" Then
                pSettings(z).RevisionVer = pv(1)
            ElseIf pv(0) = "AutoIncrementVer" Then
                pSettings(z).AutoIncrementVer = pv(1)
            ElseIf pv(0) = "ServerSupportFiles" Then
                pSettings(z).ServerSupportFiles = pv(1)
            ElseIf pv(0) = "VersionComments" Then
                pSettings(z).VersionComments = pv(1)
            ElseIf pv(0) = "VersionCompanyName" Then
                pSettings(z).VersionCompanyName = pv(1)
            ElseIf pv(0) = "VersionFileDescription" Then
                pSettings(z).VersionFileDescription = pv(1)
            ElseIf pv(0) = "VersionLegalCopyright" Then
                pSettings(z).VersionLegalCopyright = pv(1)
            ElseIf pv(0) = "VersionLegalTrademarks" Then
                pSettings(z).VersionLegalTrademarks = pv(1)
            ElseIf pv(0) = "VersionProductName" Then
                pSettings(z).VersionProductName = pv(1)
            ElseIf pv(0) = "CompilationType" Then
                pSettings(z).CompilationType = pv(1)
            ElseIf pv(0) = "OptimizationType" Then
                pSettings(z).OptimizationType = pv(1)
            ElseIf pv(0) = "FavorPentiumPro(tm)" Then
                pSettings(z).FavorPentiumPro = pv(1)
            ElseIf pv(0) = "CodeViewDebugInfo" Then
                pSettings(z).CodeViewDebugInfo = pv(1)
            ElseIf pv(0) = "NoAliasing" Then
                pSettings(z).NoAliasing = pv(1)
            ElseIf pv(0) = "BoundsCheck" Then
                pSettings(z).BoundsCheck = pv(1)
            ElseIf pv(0) = "OverflowCheck" Then
                pSettings(z).OverflowCheck = pv(1)
            ElseIf pv(0) = "FlPointCheck" Then
                pSettings(z).FlPointCheck = pv(1)
            ElseIf pv(0) = "FDIVCheck" Then
                pSettings(z).FDIVCheck = pv(1)
            ElseIf pv(0) = "UnroundedFP" Then
                pSettings(z).UnroundedFP = pv(1)
            ElseIf pv(0) = "StartMode" Then
                pSettings(z).StartMode = pv(1)
            ElseIf pv(0) = "Unattended" Then
                pSettings(z).Unattended = pv(1)
            ElseIf pv(0) = "Retained" Then
                pSettings(z).Retained = pv(1)
            ElseIf pv(0) = "ThreadPerObject" Then
                pSettings(z).ThreadPerObject = pv(1)
            ElseIf pv(0) = "MaxNumberOfThreads" Then
                pSettings(z).MaxNumberOfThreads = pv(1)
            ElseIf pv(0) = "DebugStartupOption" Then
                pSettings(z).DebugStartupOption = pv(1)
                'ElseIf pv(0) = "AutoRefresh" Then
                'pSettings(z).AutoRefresh = pv(1)
            Else
                ReDim Preserve pSettings(z).Unknown(UBound(pSettings(z).Unknown) + 1)
                pSettings(z).Unknown(UBound(pSettings(z).Unknown)) = pv(1)
            End If
        Else
            '            If inp.IndexOf("[") = 0 AndAlso inp.LastIndexOf("]") = inp.Length - 1 Then
            '              'a command or task
            '            End If
            ReDim Preserve pSettings(z).Unknown(UBound(pSettings(z).Unknown) + 1)
            pSettings(z).Unknown(UBound(pSettings(z).Unknown)) = inp
        End If
    Loop
    sr.Close
End Sub
Friend Sub SaveProjects()
    On Error Resume Next
    If vbEnv.VBProjects.Count = 0 Then
        MsgBox "You need an active project to save components"
        Exit Sub
    End If
    vbEnv.CommandBars("Standard").Controls("Sa&ve Projects").enabled = False
    Dim h1 As Long
    h1 = apiLoadCursor(ByVal 0&, IDC_APPSTARTING)
    apiSetSystemCursor h1, OCR_NORMAL  'replaces the contents of the system cursor specified by OCR_NORMAL with the contents of the 'size-all' cursor
    Dim vps As New VBIDE.VBProjects
    Set vps = vbEnv.VBProjects
    Dim i As Long
    For i = 1 To vps.Count
        SaveAllProjectFiles vps.Item(i)
    Next
    apc.ApcProject.Dirty = False
    isDirty = False
    Set vps = Nothing
    apiSetSystemCursor h1, OCR_NORMAL  'replaces the contents of the system cursor specified by OCR_NORMAL with the contents of the 'size-all' cursor
    apiDestroyCursor h1
    vbEnv.CommandBars("Standard").Controls("Sa&ve Projects").enabled = True
End Sub
Friend Sub SaveAllProjectFiles(ByRef vbproj As VBIDE.VBProject)
    On Error Resume Next
    savingall = True
    DoEvents
    apiSleep 400
    If vbproj Is Nothing Then Set vbproj = vbEnv.ActiveVBProject
    If vbproj Is Nothing Then Exit Sub
    If UBound(pSettings) = 0 Then Exit Sub
    Dim i   As Long
    Dim fso As Object
    Dim sw  As Object
    Dim pd  As String
    pd = GetProjectDirectory(vbproj)
    If Dir(pd, vbDirectory) = "" Then MkDir pd 'make the directory if it does not already exist
    Set fso = CreateObject("Scripting.FileSystemObject") 'Save the project file
    Dim app As String
    app = GetProjectPath(vbproj)
    Dim z As Long
    z = 1
    For i = 1 To UBound(pSettings)
        If pSettings(i).Name = vbproj.Name Then
            z = i
            Exit For
        End If
    Next
    Set sw = fso.CreateTextFile(app, True)
    sw.WriteLine ("Type=" & pSettings(z).ProjectType)
    Dim ext    As String
    Dim n      As String
    Dim pth    As String
    Dim refs() As String
    refs = GetListOfReferences
    If UBound(refs) > 0 Then
        For i = 1 To UBound(refs)
            sw.WriteLine ("Reference=" & refs(i))
        Next
    End If
    If UBound(pSettings(z).Objects) > 0 Then
        For i = 1 To UBound(pSettings(z).Objects)
            sw.WriteLine ("Object=" & pSettings(z).Objects(i))
        Next
    End If
    If UBound(pSettings(z).UserControls) > 0 Then
        For i = 1 To UBound(pSettings(z).UserControls)
            sw.WriteLine ("UserControl=" & pSettings(z).UserControls(i))
        Next
    End If
    If vbproj.VBComponents.Count > 0 Then
        For i = 1 To vbproj.VBComponents.Count
            ext = GetFileExtension(vbproj.VBComponents.Item(i))
            If ext = ".frm" Then
                n = vbproj.VBComponents.Item(i).Name
                pth = pd & "\" & n & ext
                sw.WriteLine ("Designer=" & pth)
                vbproj.VBComponents.Item(i).Export (pth)
            ElseIf ext = ".bas" Then
                n = vbproj.VBComponents.Item(i).Name
                pth = pd & "\" & n & ext
                sw.WriteLine ("Module=" & n & "; " & pth)
            ElseIf ext = ".cls" Then
                n = vbproj.VBComponents.Item(i).Name
                pth = pd & "\" & n & ext
                sw.WriteLine ("Class=" & n & "; " & pth)
            End If
        Next
    End If
    If UBound(pSettings(z).ResFile32) > 0 Then
        For i = 1 To UBound(pSettings(z).ResFile32)
            sw.WriteLine ("ResFile32=" & pSettings(z).ResFile32(i))
        Next
    End If
    If UBound(pSettings(z).RelatedDoc) > 0 Then
        For i = 1 To UBound(pSettings(z).RelatedDoc)
            sw.WriteLine ("RelatedDoc=" & pSettings(z).RelatedDoc(i))
        Next
    End If
    sw.WriteLine ("IconForm=" & Chr(34) & Replace(pSettings(z).IconForm, Chr(34), "") & Chr(34))
    sw.WriteLine ("Startup=" & pSettings(z).Startup)
    sw.WriteLine ("HelpFile=" & Chr(34) & Replace(vbproj.HelpFile, Chr(34), "") & Chr(34))
    sw.WriteLine ("Title=" & Chr(34) & Replace(pSettings(z).title, Chr(34), "") & Chr(34))
    sw.WriteLine ("ExeName32=" & Chr(34) & Replace(vbproj.Name & ".exe", Chr(34), "") & Chr(34))
    sw.WriteLine ("Path32=" & pSettings(z).Path32)
    sw.WriteLine ("Command32=" & Chr(34) & Replace(pSettings(z).Command32, Chr(34), "") & Chr(34))
    sw.WriteLine ("Name=" & Chr(34) & Replace(vbproj.Name, Chr(34), "") & Chr(34))
    sw.WriteLine ("HelpContextID=" & Chr(34) & Replace(pSettings(z).HelpContextID, Chr(34), "") & Chr(34))
    sw.WriteLine ("Description=" & Chr(34) & Replace(vbproj.Description, Chr(34), "") & Chr(34))
    sw.WriteLine ("CompatibleMode=" & pSettings(z).CompatibleMode)
    sw.WriteLine ("MajorVer=" & pSettings(z).MajorVer)
    sw.WriteLine ("MinorVer=" & pSettings(z).MinorVer)
    sw.WriteLine ("RevisionVer=" & pSettings(z).RevisionVer)
    sw.WriteLine ("AutoIncrementVer=" & pSettings(z).AutoIncrementVer)
    sw.WriteLine ("ServerSupportFiles=" & pSettings(z).ServerSupportFiles)
    sw.WriteLine ("VersionComments=" & pSettings(z).VersionComments)
    sw.WriteLine ("VersionCompanyName=" & pSettings(z).VersionCompanyName)
    sw.WriteLine ("VersionFileDescription=" & pSettings(z).VersionFileDescription)
    sw.WriteLine ("VersionLegalCopyright=" & pSettings(z).VersionLegalCopyright)
    sw.WriteLine ("VersionLegalTrademarks=" & pSettings(z).VersionLegalTrademarks)
    sw.WriteLine ("VersionProductName=" & pSettings(z).VersionProductName)
    sw.WriteLine ("CompilationType=" & pSettings(z).CompilationType)
    sw.WriteLine ("OptimizationType=" & pSettings(z).OptimizationType)
    sw.WriteLine ("FavorPentiumPro(tm)=" & pSettings(z).FavorPentiumPro)
    sw.WriteLine ("CodeViewDebugInfo=" & pSettings(z).CodeViewDebugInfo)
    sw.WriteLine ("NoAliasing=" & pSettings(z).NoAliasing)
    sw.WriteLine ("BoundsCheck=" & pSettings(z).BoundsCheck)
    sw.WriteLine ("OverflowCheck=" & pSettings(z).OverflowCheck)
    sw.WriteLine ("FlPointCheck=" & pSettings(z).FlPointCheck)
    sw.WriteLine ("FDIVCheck=" & pSettings(z).FDIVCheck)
    sw.WriteLine ("UnroundedFP=" & pSettings(z).UnroundedFP)
    sw.WriteLine ("StartMode=" & pSettings(z).StartMode)
    sw.WriteLine ("Unattended=" & pSettings(z).Unattended)
    sw.WriteLine ("Retained=" & pSettings(z).Retained)
    sw.WriteLine ("ThreadPerObject=" & pSettings(z).ThreadPerObject)
    sw.WriteLine ("MaxNumberOfThreads=" & pSettings(z).MaxNumberOfThreads)
    sw.WriteLine ("DebugStartupOption=" & pSettings(z).DebugStartupOption)
    'sw.writeline ("AutoRefresh=" & pSettings(z).AutoRefresh)
    If UBound(pSettings(z).Unknown) > 0 Then
        For i = 1 To UBound(pSettings(z).Unknown)
            If Trim(pSettings(z).Unknown(i)) <> "" Then sw.WriteLine (pSettings(z).Unknown(i))
        Next
    End If
    sw.Close
    Set sw = Nothing
    Set fso = Nothing
    Dim cmname     As String
    Dim cExtension As String
    Dim cLineCount As Long
    Dim cCode      As String
    Dim cHeader    As String
    Dim mpth       As String
    For i = 1 To vbproj.VBComponents.Count   'Save code modules
        cmname = ""
        cExtension = ""
        cLineCount = 0
        cCode = ""
        cHeader = ""
        mpth = ""
        cmname = vbproj.VBComponents.Item(i).Name
        cExtension = GetFileExtension(vbproj.VBComponents.Item(i))
        cLineCount = vbproj.VBComponents.Item(i).CodeModule.CountOfLines
        If cLineCount > 0 Then cCode = vbproj.VBComponents.Item(i).CodeModule.Lines(1, cLineCount)
        If cExtension = ".frm" Then
            cHeader = GetDefaultFormHeader(cmname)
        ElseIf cExtension = ".bas" Then
            cHeader = GetDefaultModuleHeader(cmname)
        ElseIf cExtension = ".cls" Then
            cHeader = GetDefaultClassHeader(cmname)
        Else 'use the default VBA header, ie nothing at all
            cHeader = ""
        End If
        If cExtension <> ".frm" Then
            If cHeader <> "" Then cCode = cHeader & vbCrLf & cCode
            mpth = pd & "\" & cmname & cExtension
            If Dir(mpth, vbNormal) <> "" Then Kill mpth
            Set fso = CreateObject("Scripting.FileSystemObject")
            Set sw = fso.CreateTextFile(mpth, True)
            sw.WriteLine cCode
            sw.Close
            Set sw = Nothing
            Set fso = Nothing
        End If
    Next
    Set sw = Nothing
    Set fso = Nothing
    Dim vbp As MSAPC.Project
    For Each vbp In m_apcHost.Projects
        vbp.Dirty = False
    Next
    Set vbp = Nothing
    isDirty = False
End Sub
Friend Sub SyncActiveProject()
    On Error Resume Next
    Dim vbproj As MSAPC.Project
    Dim dirt   As Boolean
    dirt = False
    For Each vbproj In m_apcHost.Projects
        If vbproj.Dirty = True Then dirt = True
        If vbproj.Name = vbEnv.ActiveVBProject.Name Then
            Set m_ApcProject = vbproj
        End If
    Next
    Set vbproj = Nothing
    isDirty = dirt
End Sub
Friend Function CancelDirtyCheck(Optional prjName As String = "") As Boolean
    On Error Resume Next
    CancelDirtyCheck = False
    SyncActiveProject
    If isDirty = True Then
        frmClose.Tag = prjName
        frmClose.Show vbModal, frmVB6x
        If cancelclose = True Then
            cancelclose = False
            CancelDirtyCheck = True
        End If
    End If
End Function
Private Sub SleepEvents(ByVal dwMilliseconds As Long)
    Dim i As Long
    Dim q As Long
    q = dwMilliseconds / 400
    For i = 1 To q
        DoEvents
        apiSleep 400
    Next
End Sub
Friend Sub MakeExe(Optional ByVal silent As Boolean = False, Optional ByVal vbpro As VBProject = Nothing)
    On Error GoTo skip
    If vbpro Is Nothing Then Set vbpro = vbEnv.ActiveVBProject
    If vbpro.VBComponents.Count = 0 Then
        MsgBox "No code modules assosiated with project " & vbpro.Name
        Exit Sub
    End If
    If UBound(pSettings) = 0 Then
        MsgBox "No project settings attached to project " & vbpro.Name
        Exit Sub
    End If
    Dim k        As Long
    Dim t        As Long
    Dim hasstart As Boolean
    Dim hassub   As Boolean
    For k = 1 To UBound(pSettings)
        If pSettings(k).Name = vbpro.Name Then
            For t = 1 To vbpro.VBComponents.Count
                If Replace(pSettings(k).Startup, Chr(34), "") = vbpro.VBComponents.Item(t).Name Then
                    hasstart = True
                End If
            Next
        End If
    Next
    Dim prs() As String
    Dim pr    As Variant
    For k = 1 To vbpro.VBComponents.Count
        If vbpro.VBComponents.Item(k).Type = vbext_ct_StdModule Then
            prs = ListProcedures(vbpro.VBComponents.Item(k).CodeModule)
            If UBound(prs) > 0 Then
               For Each pr In prs
                   If LCase(pr) = "main" Then
                       hassub = True
                       Exit For
                   End If
               Next
            End If
        End If
    Next

      
    If hasstart = False And hassub = False Then
       MsgBox "No startup module set to project"
        Exit Sub
    End If

    
    Dim h1 As Long
    vbEnv.CommandBars("Standard").Controls("Make .exe...").enabled = False
    vbEnv.CommandBars("Standard").Controls("Sa&ve Projects").enabled = False
    DoEvents
    Dim errlevel As Long
    errlevel = 4
    Dim fso  As Object
    Dim pth2 As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    errlevel = 3
    If vbpro Is Nothing Then GoTo skip
    errlevel = 2
    h1 = apiLoadCursor(ByVal 0&, IDC_APPSTARTING)
    apiSetSystemCursor h1, OCR_NORMAL
    Dim vbproj  As MSAPC.Project
    Dim vbprojs As MSAPC.Projects
    Set vbprojs = apc.ApcHost.Projects
    For Each vbproj In vbprojs
        If vbproj.Name = vbpro.Name Then
            If vbproj.Compile = False Then
                Set vbprojs = Nothing
                GoTo skip
            End If
            Exit For
        End If
    Next
    Set vbprojs = Nothing
    errlevel = 1
    SaveAllProjectFiles vbpro
    SleepEvents 400
    Dim csd As String
    Dim pp  As String
    Dim apd As String
    pth2 = app.Path & "\errors.txt"
    csd = GetCurrentDirectory()
    apd = GetProjectDirectory(vbpro)
    pp = GetProjectPath(vbpro)
    Call apiSetCurrentDirectory(apd)
    If fso.fileexists(pth2) = True Then Kill pth2
    Dim ret  As Long
    Dim cmdl As String
    cmdl = Chr(34) & pthVB6EXE & Chr(34) & " /make " & Chr(34) & pp & Chr(34) & " /out " & Chr(34) & pth2 & Chr(34) & " /outdir " & Chr(34) & apd & Chr(34)
    ret = apiWinExec(cmdl, 1)
    Call apiSetCurrentDirectory(csd)
    apiSetSystemCursor h1, OCR_NORMAL
    If ERROR_SUCCESS Then
        Dim pth As String
        pth = apd & "\" & vbpro.Name & ".exe"
        If silent = False Or Dir(pth, vbNormal) = "" Then
            apiSetSystemCursor h1, OCR_NORMAL
            Dim z As Long
            Do
                z = z + 1
                If z = 15 Or z Mod 45 = 0 Then
                    Dim dres As VbMsgBoxResult
                    dres = MsgBox("The compile process is taking longer than expected." & vbCrLf & "Would you like to continue compiling?", vbYesNo, "WARNING")
                    If dres = vbNo Then Exit Do
                End If
                DoEvents
                apiSleep 400
                If Dir(pth, vbNormal) <> "" And Dir(pth2, vbNormal) <> "" Then
                    SleepEvents 400
                    Exit Do
                End If
            Loop
            apiSetSystemCursor h1, OCR_NORMAL
            Set fso = CreateObject("Scripting.FileSystemObject")
            If fso.fileexists(pth) = False Then
                vbEnv.CommandBars("Standard").Controls("Make .exe...").enabled = True
                vbEnv.CommandBars("Standard").Controls("Sa&ve Projects").enabled = True
                UpdateSyntaxErrorDisplay "Executable file was never created"
                Set fso = Nothing
                Exit Sub
            End If
            If fso.fileexists(pth2) = False Then
                vbEnv.CommandBars("Standard").Controls("Make .exe...").enabled = True
                vbEnv.CommandBars("Standard").Controls("Sa&ve Projects").enabled = True
                 UpdateSyntaxErrorDisplay "Build status was never created"
                Set fso = Nothing
                Exit Sub
            End If
            Dim sr  As Object
            Dim txt As String
            Dim b   As Boolean
            Dim j   As Long
            apiSetSystemCursor h1, OCR_NORMAL
            For j = 1 To 10
                Set sr = fso.OpenTextFile(pth2, 1, False, 0)
                Do
                    If sr.AtEndOfStream = True Then Exit Do
                    Dim inp As String
                    inp = Trim(sr.Readline)
                    txt = txt & inp & vbCrLf
                    b = True
                Loop
                sr.Close
                If Trim(Replace(txt, vbCrLf, "")) <> "" Then Exit For
                SleepEvents 400
            Next
            apiSetSystemCursor h1, OCR_NORMAL
            Set fso = Nothing
            If Trim(Replace(txt, vbCrLf, "")) = "" Then txt = "Error for no reason, the PC felt like it" ' Err.Description
            If Dir(pth, vbNormal) = "" Then
                UpdateSyntaxErrorDisplay txt & vbCrLf & vbCrLf & "The executable file does not exist in path: " & pth
                'fire off internal pre-compiler to throw a specific error and select text with problem
                vbEnv.CommandBars.FindControl(msoControlButton, 578).Execute
            Else
               ' MsgBox Trim(txt)
                UpdateSyntaxErrorDisplay Trim(txt)
            End If
        End If
    ElseIf ret = ERROR_OUT_OF_MEM Then
                UpdateSyntaxErrorDisplay "Out of memory..."
    ElseIf ERROR_BAD_FORMAT Then
        UpdateSyntaxErrorDisplay "bad format.."
    ElseIf ERROR_FILE_NOT_FOUND Then
        UpdateSyntaxErrorDisplay "File not found..."
    ElseIf ERROR_PATH_NOT_FOUND Then
        UpdateSyntaxErrorDisplay "Path not found..."
    ElseIf ERROR_NO_ASSOC Then
        UpdateSyntaxErrorDisplay "No file association..."
    Else
        UpdateSyntaxErrorDisplay "Unknown error.  Perhaps, an access violation without proper rights to write to disk"
    End If
    vbEnv.CommandBars("Standard").Controls("Make .exe...").enabled = True
    vbEnv.CommandBars("Standard").Controls("Sa&ve Projects").enabled = True
    Exit Sub
skip:
    apiSetSystemCursor h1, OCR_NORMAL
    If errlevel = 1 Then UpdateSyntaxErrorDisplay "Error automating compile with VB6 - " & Err.Description
    If errlevel = 2 Then vbEnv.CommandBars.FindControl(msoControlButton, 578).Execute
    If errlevel = 3 Then UpdateSyntaxErrorDisplay "Error getting an active project to work with - " & Err.Description
    If errlevel = 4 Then UpdateSyntaxErrorDisplay "Error deleting error file inside the path of VB65 - " & Err.Description
    vbEnv.CommandBars("Standard").Controls("Make .exe...").enabled = True
    vbEnv.CommandBars("Standard").Controls("Sa&ve Projects").enabled = True
End Sub
Private Function GetListOfReferences() As String()
    On Error Resume Next
    Dim r() As String
    GetListOfReferences = r
    If vbEnv.ActiveVBProject Is Nothing Then Exit Function
    Dim rnam   As String
    Dim rdes   As String
    Dim rgui   As String
    Dim rpth   As String
    Dim rmaj   As String
    Dim rmin   As String
    Dim rPro   As String
    Dim rtyp   As String
    Dim rcount As Long
    Dim i      As Long
    For i = 1 To vbEnv.ActiveVBProject.References.Count
        Dim ri As Object
        Set ri = vbEnv.ActiveVBProject.References.Item(i)
        rgui = ri.guid
        If Trim(rgui) <> "" Then 'must have a guid value
            If rgui <> "{000204EF-0000-0000-C000-000000000046}" Then
                rpth = ri.fullpath
                If rpth <> app.Path & "\" & app.EXEName & ".exe" Then
                    With ri
                        rnam = .Name
                        rdes = .Description
                        rmaj = .Major
                        rmin = .Minor
                        rtyp = .Type
                    End With
                    If ri.IsBroken = True Then
                        ' TODO
                        MsgBox "Reference " & ri.Name & "  is missing or broken, please add the correct reference."
                        
                    End If
                    If rmaj = "" Then rmaj = "0"
                    If rmin = "" Then rmin = "0"
                    rmaj = "#" & rmaj
                    If Trim(rpth) <> "" Then rpth = "#" & rpth
                    If Trim(rdes) <> "" Then rdes = "#" & rdes
                    If Trim(rnam) <> "" Then rnam = "#" & rnam
                    If Trim(rdes) = "" Then rdes = rnam
                    rPro = "*\G" & rgui & rmaj & "." & rmin & "#0" & rpth & rdes
                    rcount = rcount + 1
                    ReDim Preserve r(rcount)
                    r(rcount) = rPro
                End If
            End If
        End If
    Next
    GetListOfReferences = r
End Function
Friend Function GetProjectDirectory(ByRef vbproj As VBIDE.VBProject) As String
    On Error Resume Next
    If vbproj Is Nothing Then Exit Function
    GetProjectDirectory = app.Path & "\Projects\" & vbproj.Name
    If Dir(GetProjectDirectory, vbDirectory) = "" Then MkDir GetProjectDirectory
End Function
Friend Function GetProjectPath(ByRef vbproj As VBIDE.VBProject) As String
    On Error Resume Next
    If vbproj Is Nothing Then Exit Function
    GetProjectPath = app.Path & "\Projects\" & vbproj.Name & "\" & vbproj.Name & ".vbp"
End Function
Friend Function GetDefaultClassHeader(ByVal fn As String) As String
    Dim txt As String
    txt = txt & "VERSION 1.0 CLASS" & vbCrLf
    txt = txt & "BEGIN" & vbCrLf
    txt = txt & "   MultiUse = -1" & vbCrLf
    txt = txt & "   Persistable = 0" & vbCrLf
    txt = txt & "   DataBindingBehavior = 0" & vbCrLf
    txt = txt & "   DataSourceBehavior = 0" & vbCrLf
    txt = txt & "   MTSTransactionMode = 0" & vbCrLf
    txt = txt & "END" & vbCrLf
    txt = txt & "Attribute VB_Name = " & Chr(34) & fn & Chr(34) & vbCrLf
    txt = txt & "Attribute VB_GlobalNameSpace = False" & vbCrLf
    txt = txt & "Attribute VB_Creatable = True" & vbCrLf
    txt = txt & "Attribute VB_PredeclaredId = False" & vbCrLf
    txt = txt & "Attribute VB_Exposed = False" & vbCrLf
    GetDefaultClassHeader = txt
End Function
Friend Function GetDefaultModuleHeader(ByVal fn As String) As String
    Dim txt As String
    txt = txt & "Attribute VB_Name = " & Chr(34) & fn & Chr(34)
    GetDefaultModuleHeader = txt
End Function
Friend Function GetDefaultFormHeader(ByVal fname As String) As String
    Dim txt As String
    txt = txt & "VERSION 5.00" & vbCrLf
    txt = txt & "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} " & fname & vbCrLf
    txt = txt & "   Caption         =   " & Chr(34) & fname & Chr(34) & vbCrLf
    txt = txt & "   ClientHeight    =   3180" & vbCrLf
    txt = txt & "   ClientLeft      =   45" & vbCrLf
    txt = txt & "   ClientTop       =   375" & vbCrLf
    txt = txt & "   ClientWidth     =   4710" & vbCrLf
    txt = txt & "   OleObjectBlob   =   " & Chr(34) & fname & ".frx" & Chr(34) & ":0000" & vbCrLf
    txt = txt & "   StartUpPosition =   1  'CenterOwner" & vbCrLf
    txt = txt & "End" & vbCrLf
    txt = txt & "Attribute VB_Name = " & Chr(34) & fname & Chr(34) & vbCrLf
    txt = txt & "Attribute VB_GlobalNameSpace = False" & vbCrLf
    txt = txt & "Attribute VB_Creatable = False" & vbCrLf
    txt = txt & "Attribute VB_PredeclaredId = True" & vbCrLf
    txt = txt & "Attribute VB_Exposed = False" & vbCrLf
    GetDefaultFormHeader = txt
End Function
Friend Function GetFileExtension(vbcomp As VBIDE.VBComponent) As String
    Select Case vbcomp.Type
        Case vbext_ct_StdModule
            GetFileExtension = ".bas"
        Case vbext_ct_ClassModule
            GetFileExtension = ".cls"
        Case vbext_ct_MSForm
            GetFileExtension = ".frm"
        Case vbext_ct_Document
            GetFileExtension = ".doc"
        Case Else
            GetFileExtension = "." & CStr(vbcomp.Type)
    End Select
End Function
Friend Sub DeleteModuleByName(ByVal Name As String)
    On Error Resume Next
    Dim vbproj As New VBIDE.VBProject
    Set vbproj = vbEnv.ActiveVBProject
    If vbproj Is Nothing Then Exit Sub
    Dim vbcomp As New VBIDE.VBComponent
    Set vbcomp = vbproj.VBComponents(Name)
    If Not vbcomp Is Nothing Then vbproj.VBComponents.Remove vbcomp
    Set vbcomp = Nothing
    Set vbproj = Nothing
End Sub
Friend Sub RemoveActiveProjectCode()
    On Error Resume Next
    Dim vbproj As New VBIDE.VBProject
    Set vbproj = vbEnv.ActiveVBProject
    If vbproj Is Nothing Then Exit Sub
    Dim vbcomp   As New VBIDE.VBComponent
    Dim vbcodmod As New VBIDE.CodeModule
    For Each vbcomp In vbproj.VBComponents
        If vbcomp.Type = vbext_ct_Document Then
            Set vbcodmod = vbcomp.CodeModule
            If Not vbcodmod Is Nothing Then
                With vbcodmod
                    .DeleteLines 1, .CountOfLines
                End With
            End If
            Set vbcodmod = Nothing
        Else
            vbproj.VBComponents.Remove vbcomp
        End If
    Next
    Set vbcodmod = Nothing
    Set vbcomp = Nothing
    Set vbproj = Nothing
End Sub
Friend Sub RenameComponent(ByVal OldName As String, ByVal NewName As String)
    On Error Resume Next
    If Trim(NewName) = "" Then Exit Sub
    If vbEnv.ActiveVBProject Is Nothing Then Exit Sub
    vbEnv.ActiveVBProject.VBComponents(OldName).Name = NewName
End Sub
Friend Function ComponentTypeToString(ComponentType As VBIDE.vbext_ComponentType) As String
    Select Case ComponentType
        Case vbext_ct_ActiveXDesigner
            ComponentTypeToString = "ActiveX Designer"
        Case vbext_ct_ClassModule
            ComponentTypeToString = "Class Module"
        Case vbext_ct_Document
            ComponentTypeToString = "Document Module"
        Case vbext_ct_MSForm
            ComponentTypeToString = "UserForm"
        Case vbext_ct_StdModule
            ComponentTypeToString = "Code Module"
        Case Else
            ComponentTypeToString = "Unknown Type: " & CStr(ComponentType)
    End Select
End Function
Friend Function ListProjects() As String
    On Error Resume Next
    Dim txt     As String
    Dim vbprojs As MSAPC.Projects
    Set vbprojs = m_apcHost.Projects
    If Not vbprojs Is Nothing Then
        Dim vbproj As MSAPC.Project
        For Each vbproj In vbprojs
            txt = txt & vbproj.Name & " (" & vbproj.VBProject.VBComponents.Count & " code module/s)" & vbCrLf
            txt = txt & ListComponents(vbproj) & vbCrLf
        Next
        Set vbprojs = Nothing
        Set vbproj = Nothing
    End If
    ListProjects = txt
End Function
Friend Function ListComponents(ByVal vbproj As MSAPC.Project) As String
    On Error Resume Next
    Dim txt As String
    If Not vbproj Is Nothing Then
        Dim vbproj2 As VBIDE.VBProject
        Set vbproj2 = vbproj.VBProject
        Dim procs() As String
        Dim proc    As Variant
        If Not vbproj2 Is Nothing Then
            Dim vbcomp As VBIDE.VBComponent
            For Each vbcomp In vbproj2.VBComponents
                procs = ListProcedures(vbcomp.CodeModule)
                Dim pcnt As Long
                pcnt = UBound(procs)
                txt = txt & "       " & vbcomp.Name & " (" & CStr(pcnt) & " procedure/s)" & vbCrLf  '& "   " & ComponentTypeToString(vbcomp.Type) & " " & vbcomp.Name & vbCrLf
                Dim j    As Long
                Dim vs() As String
                Dim vc   As Long
                Dim v    As Variant
                If pcnt > 0 Then
                    For j = 1 To pcnt
                        vs = ListVariables(vbcomp.CodeModule, procs(j))
                        vc = UBound(vs)
                        If Trim(procs(j)) <> "" Then txt = txt & "           " & procs(j) & " (" & CStr(vc) & " variable/s)" & vbCrLf
'                        For Each v In vs
'                            If Trim(v) <> "" Then txt = txt & "              " & v & vbCrLf
'                        Next
                    Next
                End If
            Next
            Set vbproj2 = Nothing
            Set vbcomp = Nothing
        End If
        Set vbproj = Nothing
    End If
    ListComponents = txt
End Function
Friend Function VariablesCount(ByVal vbcodmod As VBIDE.CodeModule, ByVal proc As String) As Long
    Dim i1 As Long
    i1 = vbcodmod.ProcStartLine(proc, vbext_pk_Proc)
    Dim i2 As Long
    i2 = i2 + vbcodmod.ProcCountLines(proc, vbext_pk_Proc)
    Dim i As Long
    For i = i1 To i2
        Dim l As String
        l = Trim(LCase(vbcodmod.Lines(i, 1)))
        If Left(l, 4) = "dim " Then VariablesCount = VariablesCount + 1
    Next
End Function
Friend Function ListVariables(ByVal vbcodmod As VBIDE.CodeModule, ByVal proc As String) As String()
    On Error Resume Next
    Dim vcount As Long
    Dim i      As Long
    Dim i1     As Long
    Dim i2     As Long
    Dim vars() As String
    Dim pline  As String
    i1 = vbcodmod.ProcStartLine(proc, vbext_pk_Proc)
    i2 = i1 + vbcodmod.ProcCountLines(proc, vbext_pk_Proc)
    For i = i1 To i2
        pline = Trim(vbcodmod.Lines(i, 1))
        If InStr(1, LCase(pline), "dim ") <> 0 Then
            If Left(LCase(pline), 4) = "dim " Then
                vcount = vcount + 1
                ReDim Preserve vars(vcount)
                vars(vcount) = Trim(Replace(pline, vbCrLf, ""))
            End If
        End If
    Next
    If UBound(vars) = 0 Then ReDim vars(0)
    ListVariables = vars
    Set vbcodmod = Nothing
End Function
Friend Function GetProjectFromCodeModule(ByRef vbcodmod As CodeModule) As VBProject
    If vbcodmod Is Nothing Then
       Set GetProjectFromCodeModule = Nothing
       Exit Function
    End If
    On Error Resume Next
    Dim i As Long
    Dim j As Long
    For i = 1 To vbEnv.VBProjects.Count
        For j = 1 To vbEnv.VBProjects.Item(i).VBComponents.Count
            If vbcodmod Is vbEnv.VBProjects.Item(i).VBComponents.Item(j).CodeModule Then
                Set GetProjectFromCodeModule = vbEnv.VBProjects.Item(i)
                Exit Function
            End If
        Next
    Next
    Set GetProjectFromCodeModule = Nothing
End Function
Friend Function ListProcedures(ByVal vbcodmod As VBIDE.CodeModule) As String()
    On Error Resume Next
    Dim lp() As String
    ReDim lp(0)
    If Not vbcodmod Is Nothing Then
        Dim procs  As Procedures
        Dim i      As Long
        Dim vbproj As VBProject
        Set vbproj = GetProjectFromCodeModule(vbcodmod)
        If Not vbproj Is Nothing Then
            Set procs = apc.ApcHost.Projects(vbproj.Name).ProjectItems(vbcodmod.Name).Procedures
            
            For i = 1 To procs.Count
                ReDim Preserve lp(i)
                lp(i) = procs.Item(i).Name
            Next
        End If
    End If
    ListProcedures = lp
End Function
Friend Function ProcKindString(ProcKind As vbext_ProcKind, ByVal procn As String) As String
    Select Case ProcKind
        Case vbext_pk_Get
            ProcKindString = "Property Get"
        Case vbext_pk_Let
            ProcKindString = "Property Let"
        Case vbext_pk_Set
            ProcKindString = "Property Set"
        Case vbext_pk_Proc
            If InStr(1, LCase(procn), "function ") <> 0 Then
                ProcKindString = "Function Procedure"
            ElseIf InStr(1, LCase(procn), "sub ") <> 0 Then
                ProcKindString = "Sub Procedure"
            Else
                ProcKindString = "Procedure"
            End If
        Case Else
            ProcKindString = "" ' "Unknown Type: " & CStr(ProcKind)
    End Select
End Function
Friend Function ProcedureInfo(procName As String, ProcKind As VBIDE.vbext_ProcKind, vbcodmod As VBIDE.CodeModule) As ProcInfo
    On Error Resume Next
    Dim pinfo As ProcInfo
    ProcedureInfo = pinfo
    If vbcodmod Is Nothing Then Exit Function
    Dim BodyLine As Long
    BodyLine = vbcodmod.ProcStartLine(procName, ProcKind)
    If BodyLine > 0 Then
        Dim FirstLine As String
        With vbcodmod
            pinfo.procName = procName
            pinfo.ProcKind = ProcKind
            pinfo.ProcBodyLine = .ProcBodyLine(procName, ProcKind)
            pinfo.ProcCountLines = .ProcCountLines(procName, ProcKind)
            pinfo.ProcStartLine = .ProcStartLine(procName, ProcKind)
            FirstLine = .Lines(pinfo.ProcBodyLine, 1)
            If StrComp(Left(FirstLine, Len("Public")), "Public", vbBinaryCompare) = 0 Then
                pinfo.procScope = ScopePublic
            ElseIf StrComp(Left(FirstLine, Len("Private")), "Private", vbBinaryCompare) = 0 Then
                pinfo.procScope = ScopePrivate
            ElseIf StrComp(Left(FirstLine, Len("Friend")), "Friend", vbBinaryCompare) = 0 Then
                pinfo.procScope = ScopeFriend
            Else
                pinfo.procScope = ScopeDefault
            End If
            pinfo.ProcDeclaration = GetProcedureDeclaration(vbcodmod, procName, ProcKind, LineSplitKeep)
        End With
    End If
    ProcedureInfo = pinfo
End Function
Friend Function GetProcedureDeclaration(vbcodmod As VBIDE.CodeModule, procName As String, ProcKind As VBIDE.vbext_ProcKind, Optional LineSplitBehavior As LineSplits = LineSplitRemove)
    On Error Resume Next
    If vbcodmod Is Nothing Then Exit Function
    Dim LineNum     As Long
    Dim s           As String
    Dim Declaration As String
    LineNum = vbcodmod.ProcBodyLine(procName, ProcKind)
    If Err.Number <> 0 Then Exit Function
    s = vbcodmod.Lines(LineNum, 1)
    Do While Right(s, 1) = "_"
        Select Case True
            Case LineSplitBehavior = LineSplitConvert
                s = Left(s, Len(s) - 1) & vbNewLine
            Case LineSplitBehavior = LineSplitKeep
                s = s & vbNewLine
            Case LineSplitBehavior = LineSplitRemove
                s = Left(s, Len(s) - 1) & " "
        End Select
        Declaration = Declaration & s
        LineNum = LineNum + 1
        s = vbcodmod.Lines(LineNum, 1)
    Loop
    Declaration = SingleSpace(Declaration & s)
    GetProcedureDeclaration = Declaration
End Function
Private Function SingleSpace(ByVal Text As String) As String
    On Error Resume Next
    Dim Pos As String
    Pos = InStr(1, Text, Space(2), vbBinaryCompare)
    Do
        Text = Replace(Text, Space(2), Space(1))
        Pos = InStr(1, Text, Space(2), vbBinaryCompare)
        If Pos = 0 Then Exit Do
    Loop
    SingleSpace = Text
End Function
Friend Function ShowProcedureInfo(ByVal CompName As String) As String
    On Error Resume Next
    Dim txt    As String
    Dim vbproj As New VBIDE.VBProject
    Set vbproj = vbEnv.ActiveVBProject
    If Not vbproj Is Nothing Then
        Dim vbcomp   As New VBIDE.VBComponent
        Dim procName As String
        Dim ProcKind As VBIDE.vbext_ProcKind
        Dim pinfo    As ProcInfo
        procName = "ProcedureInfo"
        ProcKind = vbext_pk_Proc
        Set vbcomp = vbproj.VBComponents(CompName)
        If Not vbcomp Is Nothing Then
            Dim vbcodmod As New VBIDE.CodeModule
            Set vbcodmod = vbcomp.CodeModule
            If Not vbcodmod Is Nothing Then
                pinfo = ProcedureInfo(procName, ProcKind, vbcodmod)
                txt = txt & "ProcName: " & pinfo.procName & vbCrLf
                txt = txt & "ProcKind: " & CStr(pinfo.ProcKind) & vbCrLf
                txt = txt & "ProcStartLine: " & CStr(pinfo.ProcStartLine) & vbCrLf
                txt = txt & "ProcBodyLine: " & CStr(pinfo.ProcBodyLine) & vbCrLf
                txt = txt & "ProcCountLines: " & CStr(pinfo.ProcCountLines) & vbCrLf
                txt = txt & "ProcScope: " & CStr(pinfo.procScope) & vbCrLf
                txt = txt & "ProcDeclaration: " & pinfo.ProcDeclaration & vbCrLf
                Set vbcodmod = Nothing
            End If
            Set vbcomp = Nothing
        End If
        Set vbproj = Nothing
    End If
    ShowProcedureInfo = txt
End Function
Friend Sub SearchCodeModule(ByVal modName As String)
    On Error Resume Next
    Dim vbproj   As New VBIDE.VBProject
    Dim FindWhat As String
    Dim sl       As Long ' start line
    Dim el       As Long ' end line
    Dim sc       As Long ' start column
    Dim ec       As Long ' end column
    Dim Found    As Boolean
    Set vbproj = vbEnv.ActiveVBProject
    If Not vbproj Is Nothing Then
        Dim vbcomp As New VBIDE.VBComponent
        Set vbcomp = vbproj.VBComponents(modName)
        If Not vbcomp Is Nothing Then
            Dim vbcodmod As New VBIDE.CodeModule
            Set vbcodmod = vbcomp.CodeModule
            If Not vbcodmod Is Nothing Then
                FindWhat = "findthis"
                With vbcodmod
                    sl = 1
                    el = .CountOfLines
                    sc = 1
                    ec = 255
                    Found = .Find(Target:=FindWhat, StartLine:=sl, StartColumn:=sc, EndLine:=el, EndColumn:=ec, wholeword:=True, MatchCase:=False, patternsearch:=False)
                    Do
                        Debug.Print "Found at: Line: " & CStr(sl) & " Column: " & CStr(sc)
                        el = .CountOfLines
                        sc = ec + 1
                        ec = 255
                        Found = .Find(Target:=FindWhat, StartLine:=sl, StartColumn:=sc, EndLine:=el, EndColumn:=ec, wholeword:=True, MatchCase:=False, patternsearch:=False)
                        If Found = False Then Exit Do
                    Loop
                End With
                Set vbcodmod = Nothing
            End If
            Set vbcomp = Nothing
        End If
        Set vbproj = Nothing
    End If
End Sub
Friend Function VBComponentExists(VBCompName As String, Optional vbproj As VBIDE.VBProject = Nothing) As Boolean
    On Error Resume Next
    Dim vbp As New VBIDE.VBProject ' This returns True or False indicating whether a VBComponent named VBCompName exists in the VBProject referenced by VBProj. If VBProj is omitted, the VBProject of the Active is used.
    If vbproj Is Nothing Then
        Set vbp = vbEnv.ActiveVBProject
    Else
        Set vbp = vbproj
    End If
    If vbp Is Nothing Then Exit Function
    VBComponentExists = CBool(Len(vbp.VBComponents(VBCompName).Name))
    Set vbp = Nothing
End Function
Friend Function TotalCodeLinesInVBComponent(vbcomp As VBIDE.VBComponent) As Long
    On Error Resume Next
    If vbcomp Is Nothing Then Exit Function
    Dim n         As Long
    Dim s         As String
    Dim LineCount As Long
    If vbcomp.Collection.parent.Protection = vbext_pp_locked Then
        TotalCodeLinesInVBComponent = -1
        Exit Function
    End If
    With vbcomp.CodeModule
        For n = 1 To .CountOfLines
            s = .Lines(n, 1)
            If Trim(s) = vbNullString Then
            ElseIf Left(Trim(s), 1) = "'" Then ' blank line, skip it
            Else ' comment line, skip it
                LineCount = LineCount + 1
            End If
        Next n
    End With
    TotalCodeLinesInVBComponent = LineCount
End Function
Friend Function TotalLinesInProject(Optional vbproj As VBIDE.VBProject = Nothing) As Long
    On Error Resume Next
    Dim vbp       As New VBIDE.VBProject
    Dim vbcomp    As New VBIDE.VBComponent
    Dim LineCount As Long
    If vbproj Is Nothing Then
        Set vbp = vbEnv.ActiveVBProject
    Else
        Set vbp = vbproj
    End If
    If vbp Is Nothing Then Exit Function
    If vbp.Protection = vbext_pp_locked Then
        TotalLinesInProject = -1
        Exit Function
    End If
    For Each vbcomp In vbp.VBComponents
        LineCount = LineCount + vbcomp.CodeModule.CountOfLines
    Next vbcomp
    TotalLinesInProject = LineCount
    Set vbp = Nothing
    Set vbcomp = Nothing
End Function
Friend Function TotalWrittenCodeLinesInProject(vbproj As VBIDE.VBProject) As Long
    On Error Resume Next
    If vbproj Is Nothing Then Exit Function
    Dim vbcomp    As New VBIDE.VBComponent
    Dim LineCount As Long
    If vbproj.Protection = vbext_pp_locked Then 'vbext_ProjectProtection
        TotalWrittenCodeLinesInProject = -1
        Exit Function
    End If
    For Each vbcomp In vbproj.VBComponents
        LineCount = LineCount + TotalCodeLinesInVBComponent(vbcomp)
    Next vbcomp
    TotalWrittenCodeLinesInProject = LineCount
    Set vbcomp = Nothing
End Function
Friend Function IsEditorInSync() As Boolean
    On Error Resume Next
    If vbEnv.ActiveCodePane Is Nothing Then IsEditorInSync = False: Exit Function ' tests if VBProject selected in Project window, and the ActiveVBProject is same as VBProject associated with ActiveCodePane.
    IsEditorInSync = vbEnv.ActiveVBProject Is vbEnv.ActiveCodePane.CodeModule.parent.Collection.parent
End Function
Friend Sub SyncVBAEditor()
    On Error Resume Next
    If Not vbEnv.ActiveCodePane Is Nothing Then Set vbEnv.ActiveVBProject = vbEnv.ActiveCodePane.CodeModule.parent.Collection.parent  ' syncs editor to ActiveVBProject and VBProject containing ActiveCodePane. This makes the project that contains the ActiveCodePane the ActiveVBProject.
End Sub
Friend Function CopyModule(ModuleName As String, FromVBProject As VBIDE.VBProject, ToVBProject As VBIDE.VBProject, OverwriteExisting As Boolean) As Boolean
    On Error Resume Next
    Dim vbcomp     As New VBIDE.VBComponent
    Dim fname      As String
    Dim CompName   As String
    Dim s          As String
    Dim SlashPos   As Long
    Dim ExtPos     As Long
    Dim TempVBComp As New VBIDE.VBComponent
    If FromVBProject Is Nothing Then
        CopyModule = False
        Exit Function
    End If
    If Trim(ModuleName) = vbNullString Then
        CopyModule = False
        Exit Function
    End If
    If ToVBProject Is Nothing Then
        CopyModule = False
        Exit Function
    End If
    If FromVBProject.Protection = vbext_pp_locked Then
        CopyModule = False
        Exit Function
    End If
    If ToVBProject.Protection = vbext_pp_locked Then
        CopyModule = False
        Exit Function
    End If
    Set vbcomp = FromVBProject.VBComponents(ModuleName)
    If Err.Number <> 0 Then
        CopyModule = False
        Exit Function
    End If
    fname = Environ("Temp") & "\" & ModuleName & ".bas" ' FName is the name of the temporary file to be  used in the Export/Import code.
    If OverwriteExisting = True Then
        If Dir(fname, vbNormal + vbHidden + vbSystem) <> vbNullString Then
            Err.Clear
            Kill fname
            If Err.Number <> 0 Then
                CopyModule = False
                Exit Function
            End If
        End If
        With ToVBProject.VBComponents
            .Remove .Item(ModuleName)
        End With
    Else
        Err.Clear
        Set vbcomp = ToVBProject.VBComponents(ModuleName)
        If Err.Number <> 0 Then
            If Err.Number = 9 Then ' module doesn't exist. ignore error.
            Else
                CopyModule = False ' other error. get out with return value of False
                Exit Function
            End If
        End If
    End If
    FromVBProject.VBComponents(ModuleName).Export FileName:=fname  ' Do the Export and Import operation using FName and then Kill FName.
    SlashPos = InStrRev(fname, "\") ' Extract the module name from the export file name.
    ExtPos = InStrRev(fname, ".")
    CompName = Mid(fname, SlashPos + 1, ExtPos - SlashPos - 1)
    Set vbcomp = Nothing ' Document modules (SheetX and ThisWorkbook)   cannot be removed. So, if we are working with a document object, delete all code in that component and add the lines of FName back in to the module.
    Set vbcomp = ToVBProject.VBComponents(CompName)
    If vbcomp Is Nothing Then
        ToVBProject.VBComponents.Import FileName:=fname
    Else
        If vbcomp.Type = vbext_ct_Document Then
            Set TempVBComp = ToVBProject.VBComponents.Import(fname) ' VBComp is destination module
            With vbcomp.CodeModule  ' TempVBComp is source module
                .DeleteLines 1, .CountOfLines
                s = TempVBComp.CodeModule.Lines(1, TempVBComp.CodeModule.CountOfLines)
                .InsertLines 1, s
            End With
            On Error GoTo 0
            ToVBProject.VBComponents.Remove TempVBComp
        End If
    End If
    Kill fname
    CopyModule = True
End Function
Friend Function AddModuleToProject(ByVal vbproj As VBProject, ByVal mType As Long, ByVal mName As String) As VBComponent
    On Error Resume Next
    Set AddModuleToProject = Nothing
    If vbproj Is Nothing Then Set vbproj = vbEnv.ActiveVBProject
    If Not vbproj Is Nothing Then
        Dim vbcomp As New VBIDE.VBComponent
        Set vbcomp = vbproj.VBComponents.Add(mType)
        If Not vbcomp Is Nothing Then
           vbcomp.Name = mName
           Set AddModuleToProject = vbcomp
           Set vbcomp = Nothing
        End If
        Set vbproj = Nothing
    End If
End Function
Friend Function AddProcedureToModule(ByVal vbproj As VBProject, ByVal modName As String, ByVal procName As String, ByVal procScope As String, ByVal procType As String, ByVal procParams As String, ByVal procReturn As String) As Procedure
    On Error Resume Next
    Dim vbproc As Procedure
    If vbproj Is Nothing Then Set vbproj = vbEnv.ActiveVBProject
    If Not vbproj Is Nothing Then
        Dim vbcomp As New VBIDE.VBComponent
        Set vbcomp = vbproj.VBComponents(modName)
        If Not vbcomp Is Nothing Then
            Dim vbcodmod As New VBIDE.CodeModule
            Set vbcodmod = vbcomp.CodeModule
            If Not vbcodmod Is Nothing Then
                Dim LineNum As Long
                If procReturn <> "" Then procReturn = " As " & procReturn
                Dim getsetlet As String
                
                With vbcodmod
                    LineNum = .CountOfLines + 1
                    .InsertLines LineNum, procScope & " " & procType & " " & procName & "(" & procParams & ")" & procReturn
                    LineNum = LineNum + 1
                    '  .InsertLines LineNum, "    MsgBox " & Chr(34) & "Hello World" & Chr(34)
                    LineNum = LineNum + 1
                    If LCase(procType) = "property get" Then procType = "Property": getsetlet = "get"
                    If LCase(procType) = "property set" Then procType = "Property": getsetlet = "set"
                    If LCase(procType) = "property let" Then procType = "Property": getsetlet = "let"
                    .InsertLines LineNum, "End " & procType
                End With
                Dim procs As Procedures
                Dim i     As Long
                Set procs = apc.ApcHost.Projects(vbproj.Name).ProjectItems(vbcodmod.Name).Procedures
                frmNewProject.visible = False
                For i = 1 To procs.Count
                    If procs.Item(i).Name = procName Then
                        If procs.Item(i).Type = axProcTypePropertyGet And getsetlet = "get" Then
                        
                            Set AddProcedureToModule = procs.Item(i)
                            Exit For
                        ElseIf procs.Item(i).Type = axProcTypePropertySet And getsetlet = "set" Then
                             Set AddProcedureToModule = procs.Item(i)
                             
                            Exit For
                        ElseIf procs.Item(i).Type = axProcTypePropertyLet And getsetlet = "let" Then
                            Set AddProcedureToModule = procs.Item(i)
                            Exit For
                        Else
                        
                        'ElseIf procs.Item(i).Type = axProcTypeEventSub Then
                        
                        'ElseIf procs.Item(i).Type = axProcTypeFunction Then
                        
                        'ElseIf procs.Item(i).Type = axProcTypeSub Then
                            Set AddProcedureToModule = procs.Item(i)
                            Exit For
                        End If
                       
                    End If
                Next
                frmNewProject.visible = True
                Set vbcodmod = Nothing
            End If
            Set vbcomp = Nothing
        End If
        Set vbproj = Nothing
    End If
End Function
Friend Sub CreateEventProcedure(ByVal componentName As String, ByVal eventName As String, ByVal objectName As String)
    On Error Resume Next
    Dim vbproj As New VBIDE.VBProject
    Set vbproj = vbEnv.ActiveVBProject
    If Not vbproj Is Nothing Then
        Dim vbcomp As New VBIDE.VBComponent
        Set vbcomp = vbproj.VBComponents(componentName)
        If Not vbcomp Is Nothing Then
            Dim vbcodmod As New VBIDE.CodeModule
            Set vbcodmod = vbcomp.CodeModule
            If Not vbcodmod Is Nothing Then
                Dim LineNum As Long
                With vbcodmod
                    LineNum = .CreateEventProc(eventName, objectName)
                    LineNum = LineNum + 1
                    .InsertLines LineNum, "    MsgBox " & Chr(34) & "Hello World" & Chr(34)
                End With
                Set vbcodmod = Nothing
            End If
            Set vbcomp = Nothing
        End If
        Set vbproj = Nothing
    End If
End Sub
Friend Sub CreateProcedure(ByVal modName As String)
    On Error Resume Next
    Dim vbproj As New VBIDE.VBProject
    Set vbproj = vbEnv.ActiveVBProject
    If Not vbproj Is Nothing Then
        Dim vbcomp As New VBIDE.VBComponent
        Set vbcomp = m_apcHost.VBE.ActiveVBProject.VBComponents(modName)
        If Not vbcomp Is Nothing Then
            Dim vbcodmod As New VBIDE.CodeModule
            Set vbcodmod = vbcomp.CodeModule
            If Not vbcodmod Is Nothing Then
                Dim s       As String
                Dim LineNum As Long
                LineNum = vbcodmod.CountOfLines + 1
                s = "Sub HelloWorld()" & vbCrLf & "    MsgBox ""Hello, World""" & vbCrLf & "End Sub"
                vbcodmod.InsertLines LineNum, s
                Set vbcodmod = Nothing
            End If
            Set vbcomp = Nothing
        End If
        Set vbproj = Nothing
    End If
End Sub
'    Function ProcedureInfo(ProcName As String, ProcKind As VBIDE.vbext_ProcKind, CodeMod As VBIDE.CodeModule) As ProcInfo
'        Dim PInfo As ProcInfo
'        Dim BodyLine As Long
'        Dim Declaration As String
'        Dim FirstLine As String
'        BodyLine = CodeMod.ProcStartLine(ProcName, ProcKind)
'        If BodyLine > 0 Then
'            With CodeMod
'                PInfo.ProcName = ProcName
'                PInfo.ProcKind = ProcKind
'                PInfo.ProcBodyLine = .ProcBodyLine(ProcName, ProcKind)
'                PInfo.ProcCountLines = .ProcCountLines(ProcName, ProcKind)
'                PInfo.ProcStartLine = .ProcStartLine(ProcName, ProcKind)
'
'                FirstLine = .Lines(PInfo.ProcBodyLine, 1)
'                If StrComp(Left(FirstLine, Len("Public")), "Public", vbBinaryCompare) = 0 Then
'                    PInfo.ProcScope = ScopePublic
'                ElseIf StrComp(Left(FirstLine, Len("Private")), "Private", vbBinaryCompare) = 0 Then
'                    PInfo.ProcScope = ScopePrivate
'                ElseIf StrComp(Left(FirstLine, Len("Friend")), "Friend", vbBinaryCompare) = 0 Then
'                    PInfo.ProcScope = ScopeFriend
'                Else
'                    PInfo.ProcScope = ScopeDefault
'                End If
'                PInfo.ProcDeclaration = GetProcedureDeclaration(CodeMod, ProcName, ProcKind, LineSplitKeep)
'            End With
'        End If
'        ProcedureInfo = PInfo
'    End Function
'
'
'    Public Function GetProcedureDeclaration(CodeMod As VBIDE.CodeModule, ProcName As String, ProcKind As VBIDE.vbext_ProcKind, Optional LineSplitBehavior As LineSplits = LineSplitRemove)
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ' GetProcedureDeclaration
'    ' This return the procedure declaration of ProcName in CodeMod. The LineSplitBehavior
'    ' determines what to do with procedure declaration that span more than one line using
'    ' the "_" line continuation character. If LineSplitBehavior is LineSplitRemove, the
'    ' entire procedure declaration is converted to a single line of text. If
'    ' LineSplitBehavior is LineSplitKeep the "_" characters are retained and the
'    ' declaration is split with vbNewLine into multiple lines. If LineSplitBehavior is
'    ' LineSplitConvert, the "_" characters are removed and replaced with vbNewLine.
'    ' The function returns vbNullString if the procedure could not be found.
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        Dim LineNum As Long
'        Dim S As String
'        Dim Declaration As String
'
'        On Error Resume Next
'        LineNum = CodeMod.ProcBodyLine(ProcName, ProcKind)
'        If Err.Number <> 0 Then
'            Exit Function
'        End If
'        S = CodeMod.Lines(LineNum, 1)
'        Do While Right(S, 1) = "_"
'            Select Case True
'                Case LineSplitBehavior = LineSplitConvert
'                    S = Left(S, Len(S) - 1) & vbNewLine
'                Case LineSplitBehavior = LineSplitKeep
'                    S = S & vbNewLine
'                Case LineSplitBehavior = LineSplitRemove
'                    S = Left(S, Len(S) - 1) & " "
'            End Select
'            Declaration = Declaration & S
'            LineNum = LineNum + 1
'            S = CodeMod.Lines(LineNum, 1)
'        Loop
'        Declaration = SingleSpace(Declaration & S)
'        GetProcedureDeclaration = Declaration
'
'
'    End Function
'
'    Private Function SingleSpace(ByVal Text As String) As String
'        Dim Pos As String
'        Pos = InStr(1, Text, Space(2), vbBinaryCompare)
'        Do Until Pos = 0
'            Text = Replace(Text, Space(2), Space(1))
'            Pos = InStr(1, Text, Space(2), vbBinaryCompare)
'        Loop
'        SingleSpace = Text
'    End Function
'
''You can call the ProcedureInfo function using code like the following:
'
'    Sub ShowProcedureInfo()
'        Dim VBProj As VBIDE.VBProject
'        Dim VBComp As VBIDE.VBComponent
'        Dim CodeMod As VBIDE.CodeModule
'        Dim CompName As String
'        Dim ProcName As String
'        Dim ProcKind As VBIDE.vbext_ProcKind
'        Dim PInfo As ProcInfo
'
'        CompName = "modVBECode"
'        ProcName = "ProcedureInfo"
'        ProcKind = vbext_pk_Proc
'
'        Set VBProj = ActiveWorkbook.VBProject
'        Set VBComp = VBProj.VBComponents(CompName)
'        Set CodeMod = VBComp.CodeModule
'
'        PInfo = ProcedureInfo(ProcName, ProcKind, CodeMod)
'
'        Debug.Print "ProcName: " & PInfo.ProcName
'        Debug.Print "ProcKind: " & CStr(PInfo.ProcKind)
'        Debug.Print "ProcStartLine: " & CStr(PInfo.ProcStartLine)
'        Debug.Print "ProcBodyLine: " & CStr(PInfo.ProcBodyLine)
'        Debug.Print "ProcCountLines: " & CStr(PInfo.ProcCountLines)
'        Debug.Print "ProcScope: " & CStr(PInfo.ProcScope)
'        Debug.Print "ProcDeclaration: " & PInfo.ProcDeclaration
'    End Sub
'Friend Function InvokeElement(ByVal hWnd As Long, ByVal sText As String) As Boolean
'    On Error GoTo poop
'    InvokeElement = False
'    If hWnd = 0 Then Exit Function
'    Dim uAuto   As IUIAutomation
'    Dim el      As IUIAutomationElement
'    Dim uCond   As IUIAutomationCondition
'    Dim elName  As IUIAutomationElement
'    Dim pInvoke As IUIAutomationInvokePattern
'    Set uAuto = New CUIAutomation
'    Set uCond = uAuto.CreatePropertyCondition(UIA_NamePropertyId, sText)
'    If uCond Is Nothing Then Exit Function
'    Set el = uAuto.ElementFromHandle(ByVal hWnd)
'    If el Is Nothing Then Exit Function
'    Set elName = el.FindFirst(TreeScope_Children, uCond)
'    If elName Is Nothing Then Exit Function
'    Set pInvoke = elName.GetCurrentPattern(UIA_InvokePatternId)
'    If pInvoke Is Nothing Then Exit Function
'    pInvoke.Invoke
'    InvokeElement = True
'poop:
'End Function
'Friend Function GetUIADesktopElements() As IUIAutomationElementArray
'    On Error GoTo poop
'    Dim allChilds As IUIAutomationElementArray
'    Set GetUIADesktopElements = allChilds
'    Dim oUIAutomation As New CUIAutomation
'    Dim oUIADesktop   As IUIAutomationElement
'    Set oUIADesktop = oUIAutomation.GetRootElement
'    If oUIADesktop Is Nothing Then Exit Function
'    Set allChilds = oUIADesktop.FindAll(TreeScope_Children, oUIAutomation.CreateTrueCondition)
'    Set GetUIADesktopElements = allChilds
'poop:
'End Function
'Friend Function GetUIAWindowElements(ByVal hWnd As Long) As IUIAutomationElementArray
'    On Error GoTo poop
'    Dim allChilds As IUIAutomationElementArray
'    Set GetUIAWindowElements = allChilds
'    Dim oUIAutomation As New CUIAutomation
'    Dim oUIAWindow    As IUIAutomationElement
'    Set oUIAWindow = oUIAutomation.ElementFromHandle(hWnd)
'    If oUIAWindow Is Nothing Then Exit Function
'    Set allChilds = oUIAWindow.FindAll(TreeScope_Children, oUIAutomation.CreateTrueCondition)
'    Set GetUIAWindowElements = allChilds
'poop:
'End Function
'Friend Function WalkUIADesktopElements() As IUIAutomationElement()
'    On Error GoTo poop
'    Dim el()          As IUIAutomationElement
'    Dim oUIAutomation As New CUIAutomation
'    Dim oUIADesktop   As IUIAutomationElement
'    Dim allChilds     As IUIAutomationElementArray
'    Dim oUIElement    As IUIAutomationElement
'    Dim oTW           As IUIAutomationTreeWalker
'    Set oUIADesktop = oUIAutomation.GetRootElement
'    If oUIADesktop Is Nothing Then Exit Function
'    Set oTW = oUIAutomation.ControlViewWalker
'    Set oUIElement = oTW.GetFirstChildElement(oUIADesktop)
'    Dim i As Long
'    Do
'        If oUIElement Is Nothing Then Exit Do ' exit loop
'        On Error GoTo skip
'        ReDim Preserve el(i)
'        Set el(i) = oUIElement
'skip:
'        Set oUIElement = oTW.GetNextSiblingElement(oUIElement)
'        i = i + 1
'    Loop
'poop:
'    WalkUIADesktopElements = el
'End Function
'Friend Function ProcessIdToHandleUIA(ByVal pid As Long) As Long
'    ProcessIdToHandleUIA = 0
'    On Error GoTo poop
'    Dim el()          As IUIAutomationElement
'    Dim oUIAutomation As New CUIAutomation
'    Dim oUIADesktop   As IUIAutomationElement
'    Dim allChilds     As IUIAutomationElementArray
'    Dim oUIElement    As IUIAutomationElement
'    Dim oTW           As IUIAutomationTreeWalker
'    Set oUIADesktop = oUIAutomation.GetRootElement
'    If oUIADesktop Is Nothing Then Exit Function
'    Set oTW = oUIAutomation.ControlViewWalker
'    Set oUIElement = oTW.GetFirstChildElement(oUIADesktop)
'    ' auwlk = Windows.Automation.TreeWalker.ControlViewWalker.GetFirstChild(AutomationElement.RootElement)
'    Dim i As Long
'    For i = 1 To 1000
'        On Error GoTo skip
'        If oUIElement Is Nothing Then Exit For
'        Dim hWnd As Long
'        ' hwnd = GetUIAHandle(oUIElement)
'        '        If hwnd <> 0 And pid = GetWindowProcessId(hwnd) Then
'        '            ProcessIdToHandleUIA = hwnd: Exit Function
'        '        End If
'        '        oUIElement = oTW.GetNextSibling(oUIElement)
'skip:
'    Next
'poop:
'End Function
'Friend Function GetElementFromPoint() As IUIAutomationElement
'    '    On Error GoTo poop
'    '    Dim UI As New UIAutomationClient.CUIAutomation
'    '    Dim El As UIAutomationClient.IUIAutomationElement
'    '    Dim tag As UIAutomationClient.tagPOINT
'    '    Set GetElementFromPoint = El
'    '    Set UI = New CUIAutomation
'    '    Dim p As POINTAPI
'    '    If apiGetCursorPos(p) = 0 Then Exit Function
'    '    tag.x = p.x
'    '    tag.y = p.y
'    '    Set El = UI.ElementFromPoint(tag)
'    '    Set GetElementFromPoint = El
'    'poop:
'End Function
'    Dim i As Long
'    Dim cLine As String
'    For i = 1 To ProjectItem.VBComponent.CodeModule.CountOfLines
'       cLine = ProjectItem.VBComponent.CodeModule.Lines(i, 1)
'       ProjectItem.VBComponent.CodeModule.ReplaceLine i, Trim(cLine)
'
'    Next
''scope
'Const vbext_Private = 1
'Const vbext_Public = 2
'Const vbext_Friend = 3
''member object
'Const vbext_mt_Method = 1
'Const vbext_mt_Property = 2
'Const vbext_mt_Variable = 3
'Const vbext_mt_Event = 4
'Const vbext_mt_Const = 5
''control types cannot be created through the object model.
'Const msoControlCustom                 As Long = 0 'Custom control.
'Const msoControlEdit                   As Long = 2 'Text box.
'Const msoControlDropdown               As Long = 3 'Drop-down list.
'Const msoControlComboBox               As Long = 4 'Combo box.
'Const msoControlButtonDropdown         As Long = 5 'Drop-down button.
'Const msoControlSplitDropdown          As Long = 6 'Split drop-down list.
'Const msoControlOCXDropdown            As Long = 7 'OCX drop-down list.
'Const msoControlGenericDropdown        As Long = 8 'Generic drop-down list.
'Const msoControlGraphicDropdown        As Long = 9 'Graphic drop-down list.
'Const msoControlGraphicPopup           As Long = 11 'Graphic pop-up menu.
'Const msoControlButtonPopup            As Long = 12 'Pop-up button.
'Const msoControlSplitButtonPopup       As Long = 13 'Split button pop-up.
'Const msoControlSplitButtonMRUPopup    As Long = 14 'Most Recently Used (MRU) pop-up.
'Const msoControlLabel                  As Long = 15 'Label.
'Const msoControlExpandingGrid          As Long = 16 'Expanding grid.
'Const msoControlSplitExpandingGrid     As Long = 17 'Split expanding grid.
'Const msoControlGrid                   As Long = 18 'Grid.
'Const msoControlGauge                  As Long = 19 'Gauge control.
'Const msoControlGraphicCombo           As Long = 20 'Graphic combo box.
'Const msoControlPane                   As Long = 21 'Pane.
'Const msoControlActiveX                As Long = 22 'ActiveX control.
'Const msoControlSpinner                As Long = 23 'Spinner.
'Const msoControlLabelEx                As Long = 24 'Extended label.
'Const msoControlWorkPane               As Long = 25 'Work pane.
'Const msoControlAutoCompleteCombo      As Long = 26 'Combo box in which the first matching choice is automatically filled in as the user types.
''Fields
'Const msoCTPDockPositionLeft           As Long = 0
'Const msoCTPDockPositionTop            As Long = 1
'Const msoCTPDockPositionRight          As Long = 2
'Const msoCTPDockPositionBottom         As Long = 3
'Const msoCTPDockPositionFloating       As Long = 4
''
'Const msoBarLeft                       As Long = 0 'Command bar is docked on the left side of the application window.
'Const msoBarTop                        As Long = 1 'Command bar is docked at the top of the application window.
'Const msoBarRight                      As Long = 2 'Command bar is docked on the right side of the application window.
'Const msoBarBottom                     As Long = 3  'Command bar is docked at the bottom of the application window.
'Const msoBarFloating                   As Long = 4 'Command bar floats on top of the application window.
'Const msoBarPopup                      As Long = 5 'Command bar will be a shortcut menu.
'Const msoBarMenuBar                    As Long = 6 'Command bar will be a menu bar (Macintosh only).
''
'Const msoButtonAutomatic               As Long = 0   ' Default behavior.
'Const msoButtonCaption                 As Long = 2   ' Text only.
'Const msoButtonIcon                    As Long = 1  '  Image only.
'Const msoButtonIconAndCaption          As Long = 3  ' Image and text, with text to the right of image.
'Const msoButtonIconAndCaptionBelow     As Long = 11 '  Image with text below.
'Const msoButtonIconAndWrapCaption      As Long = 7   ' Image with text wrapped and to the right of the image.
'Const msoButtonIconAndWrapCaptionBelow As Long = 15  ' Image with text wrapped below image.
'Const msoButtonWrapCaption             As Long = 14    'Text only, centered and wrapped.
''
'Const msoButtonDown                    As Long = -1 '  Button is pressed down.
'Const msoButtonMixed                   As Long = 2   'Button is pressed down.
'Const msoButtonUp                      As Long = 0 'Button is not pressed down.
''
'Const vbext_pt_StandardExe = 0
'Const vbext_pt_ActiveXExe = 1
'Const vbext_pt_ActiveXDll = 2
'Const vbext_pt_ActiveXControl = 3
''
'Const vbext_psm_StandAlone = 0
'Const vbext_psm_OleServer = 1
''
'Const vbext_pk_Get               As Long = 3 '. A Property Get procedure.
'Const vbext_pk_Let               As Long = 1 '. A Property Let procedure.
'Const vbext_pk_Set               As Long = 2 '). A Property Set procedure.
'Const vbext_pk_Proc              As Long = 0 '). A Sub or Function procedure.
''Vbcomponent object
'Const vbext_ct_StdModule         As Long = 1
'Const vbext_ct_ClassModule       As Long = 2
'Const vbext_ct_MSForm            As Long = 3
'Const vbext_ct_ResFile           As Long = 4
'Const vbext_ct_VBForm            As Long = 5
'Const vbext_ct_VBMDIForm         As Long = 6
'Const vbext_ct_PropPage          As Long = 7
'Const vbext_ct_UserControl       As Long = 8
'Const vbext_ct_DocObject         As Long = 9
'Const vbext_ct_RelatedDocument   As Long = 10
'Const vbext_ct_ActiveXDesigner   As Long = 11
'Const vbext_ct_Document          As Long = 100 'document module
''reference object
'Const vbext_rk_TypeLib           As Long = 0
'Const vbext_rk_Project           As Long = 1
''vbproject object
'Const vbext_pt_HostProject       As Long = 100
'Const vbext_pt_StandAlone        As Long = 101
''window object
'Const vbext_wt_CodeWindow        As Long = 0
'Const vbext_wt_Designer          As Long = 1
'Const vbext_wt_Browser           As Long = 2
'Const vbext_wt_Watch             As Long = 3
'Const vbext_wt_Locals            As Long = 4
'Const vbext_wt_Immediate         As Long = 5
'Const vbext_wt_ProjectWindow     As Long = 6
'Const vbext_wt_PropertyWindow    As Long = 7
'Const vbext_wt_Find              As Long = 8
'Const vbext_wt_FindReplace       As Long = 9
'Const vbext_wt_Toolbox           As Long = 10
'Const vbext_wt_LinkedWindowFrame As Long = 11
'Const vbext_wt_MainWindow        As Long = 12
'Const vbext_wt_ToolWindow        As Long = 15
