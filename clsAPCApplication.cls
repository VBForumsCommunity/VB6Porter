VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAPCApplication"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Main host application object representing the SDI notepad exposed properties and methods."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Property Get Application() As clsAPCApplication
    Set Application = Me
End Property
Public Property Get VBE() As Object
    Set VBE = apc.ApcHost.VBE
End Property
Public Function klinkle() As String
    ' MsgBox "yup"
    klinkle = apc.ApcHost.Projects(1).DisplayName
End Function
