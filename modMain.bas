Attribute VB_Name = "modMain"
Option Explicit

Public VBInstance           As VBIDE.VBE
Public mcmpCurrentForm      As VBComponent          'current form

Public Type ctlSel
    con() As Integer
    txtname() As String
    index() As Integer
    cnt As Integer
End Type

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long '
'
'

Public Function InRunMode(VBInst As VBIDE.VBE) As Boolean
  InRunMode = (VBInst.CommandBars("File").Controls(1).Enabled = False)
End Function
