VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   7110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   _ExtentX        =   13229
   _ExtentY        =   12541
   _Version        =   393216
   Description     =   "Tab Order"
   DisplayName     =   "Tab Order"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "None"
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim newCls                              As clsSortedCollection  ' collection arry
Dim ctlNew                              As ctlSel
Dim cnt                                 As Integer              ' holds # of controls
Dim FormDisplayed                       As Boolean              ' boolean to indate if form is selected
Dim tabOrderToolbar                     As CommandBar           ' handle for main toolbar created
Dim mainTabOrder                        As CommandBarButton     ' display the tab orders
Dim tabHelp                             As CommandBarButton
Dim tabMenuPlus                         As CommandBarButton     ' increase the tab index
Dim tabMenuMinus                        As CommandBarButton     ' decrease the tab index
Dim tabMenuSet                          As CommandBarButton     ' set the tab index order
Dim tabLeft                             As CommandBarButton     ' set tap order by left alignment
Dim tabTop                              As CommandBarButton     ' set tab order by top alignment
Dim tabSelect                           As CommandBarButton     ' set tab order by selected controls
Dim tabCon                              As CommandBarButton     ' allows selecting all controls on a container
Dim vbf                                 As VBForm               ' handle for form selected
Dim txtIndex()                          As VBControl            ' control array for tab indexes
Dim tabSet                              As Boolean              ' boolean to indicate if tab toolbar has been displayed
Dim TabAdded                            As Boolean              ' boolean to indicate if sub toolbar items have been displayed
Dim tabUpdate                           As Boolean              ' boolean to indicate if update should be performed
Dim mfrmAddIn                           As New frmTab           ' used on connect
Private WithEvents HelpHandler          As CommandBarEvents     ' command bar event to load help file
Attribute HelpHandler.VB_VarHelpID = -1
Private WithEvents MenuHandler          As CommandBarEvents     ' command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1
Private WithEvents TabPlusHandler       As CommandBarEvents     ' handle event for increasing tab index
Attribute TabPlusHandler.VB_VarHelpID = -1
Private WithEvents TabMinusHandler      As CommandBarEvents     ' handle event for decreasing tab index
Attribute TabMinusHandler.VB_VarHelpID = -1
Private WithEvents TabSetHandler        As CommandBarEvents     ' handle event for setting the tab index order
Attribute TabSetHandler.VB_VarHelpID = -1
Private WithEvents TabLeftHandler       As CommandBarEvents     ' handle event for setting the tab order based on controls Left position
Attribute TabLeftHandler.VB_VarHelpID = -1
Private WithEvents TabTopHandler        As CommandBarEvents     ' handle event for setting the tab order based on controls Top position
Attribute TabTopHandler.VB_VarHelpID = -1
Private WithEvents TabSelectHandler     As CommandBarEvents     ' handle event for selected controls
Attribute TabSelectHandler.VB_VarHelpID = -1
Private WithEvents controlSel           As SelectedVBControlsEvents
Attribute controlSel.VB_VarHelpID = -1

Sub Hide()
    
    On Error Resume Next
    
    FormDisplayed = False
    mfrmAddIn.Hide
   
End Sub

Sub Show()
  
    On Error Resume Next
    
    If mfrmAddIn Is Nothing Then
        Set mfrmAddIn = New frmTab
    End If
    
    Set mfrmAddIn.VBInstance = VBInstance
    Set mfrmAddIn.Connect = Me
    FormDisplayed = True
    'mfrmAddIn.Show
    TabAdded = False
    tabSet = False
    
End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    'save the vb instance
    Set VBInstance = Application
    Set newCls = New clsSortedCollection
    
    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        Me.Show
    Else
        AddToAddInCommandBar
    End If
  
    If ConnectMode = ext_cm_AfterStartup Then
        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
            'set this to display the form on connect
            Me.Show
        End If
    End If
    
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    Dim i As Integer
    Dim S As String
    Dim ctl As VBControl
    Dim Count As Integer
    
    'delete the command bar entry
    tabOrderToolbar.Delete
    'mainTabOrder.Delete
    'tabMenuMinus.Delete
    'tabMenuPlus.Delete
    'tabMenuSet.Delete
    'tabLeft.Delete
    'tabTop.Delete
    'tabSelect.Delete
    
    Set tabOrderToolbar = Nothing
    Set mainTabOrder = Nothing
    Set tabMenuMinus = Nothing
    Set tabMenuPlus = Nothing
    Set tabMenuSet = Nothing
    Set tabLeft = Nothing
    Set tabTop = Nothing
    Set tabSelect = Nothing
    Set tabHelp = Nothing
    Set tabCon = Nothing
    
    'shut down the Add-In
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
    
    Unload mfrmAddIn
    Set mfrmAddIn = Nothing
    Unload frmTab
    Set frmTab = Nothing
    
    i = 0
    If tabSet Then
        Count = mcmpCurrentForm.Designer.VBControls.Count
        For i = Count To 1 Step -1
            Set ctl = mcmpCurrentForm.Designer.VBControls(i)
            If ctl.Properties!Name = "txtIndex" Then
                S = "txtIndex"
                cmdRemoveControl S, i
            End If
        Next i
    End If
    
    tabSet = False
    Err.Clear
    Set ctl = Nothing
    Set newCls = Nothing
    Erase ctlNew.con
    
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        'set this to display the form on connect
        Me.Show
    End If
    
End Sub

Private Sub ConHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

    'select all on container
    '
    tabUpdate = False
    'RemoveAndUpdate selStart, selCount
    tabCon.Enabled = False
    
End Sub

Private Sub controlSel_ItemAdded(ByVal VBControl As VBIDE.VBControl)
    Dim i As Integer
    Dim cnt As Integer
    Dim ctl As VBControl
    
    On Error Resume Next
    
    i = ctlNew.cnt
    For cnt = 1 To i
        If ctlNew.con(cnt) = VBControl.Properties!Caption And ctlNew.txtname(cnt) = VBControl.Properties!Name Then
            Exit Sub
        End If
    Next cnt
    If VBControl.Properties!Name <> "txtIndex" Then
        Exit Sub
    End If
    ReDim Preserve ctlNew.con(i + 1)
    ReDim Preserve ctlNew.txtname(i + 1)
    ReDim Preserve ctlNew.index(i + 1)
    ctlNew.txtname(i + 1) = VBControl.Properties!Name
    ctlNew.con(i + 1) = VBControl.Properties!Caption
    ctlNew.index(i + 1) = VBControl.Properties!tabIndex
    ctlNew.cnt = ctlNew.cnt + 1
    If Err Then
        Err.Clear
    End If
    
    Set ctl = Nothing
    
End Sub

Private Sub controlSel_ItemRemoved(ByVal VBControl As VBIDE.VBControl)
    Dim i As Integer
    
    On Error Resume Next
    i = ctlNew.cnt
    If i > 0 Then
        ReDim Preserve ctlNew.con(i - 1)
        ReDim Preserve ctlNew.txtname(i - 1)
        ReDim Preserve ctlNew.index(i - 1)
        ctlNew.cnt = ctlNew.cnt - 1
    End If
    If Err Then
        Err.Clear
    End If
    
End Sub

Private Sub HelpHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

    ShellExecute frmTab.hwnd, "open", "C:\vb6.0\taborder\Help.html", vbNullString, vbNullString, 1
    
End Sub

'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    
    If modMain.InRunMode(VBInstance) Then Exit Sub
    'get all the tabindex'(s)
    tabUpdate = True
    RemoveAndUpdate
    
End Sub

'add our toolbar items
Private Sub AddToAddInCommandBar()
    
    On Error GoTo AddToAddInCommandBarErr
    
    Set tabOrderToolbar = VBInstance.CommandBars.Add("Tab Order", msoBarFloating, , True)
    
    With tabOrderToolbar.Controls
        'add the main commandbar
        Set mainTabOrder = .Add(msoControlButton)
        mainTabOrder.Caption = "Display tab order for controls"
        Clipboard.SetData LoadResPicture(101, 0)
        mainTabOrder.PasteFace
        Set MenuHandler = VBInstance.Events.CommandBarEvents(mainTabOrder)
        
        'add the order by selected controls
        Set tabSelect = .Add(msoControlButton)
        tabSelect.Caption = "Set tab order by order of selected controls"
        Clipboard.SetData LoadResPicture(107, 0)
        tabSelect.PasteFace
        tabSelect.Enabled = False
        Set TabSelectHandler = VBInstance.Events.CommandBarEvents(tabSelect)
        
        'add the order by controls left position
        Set tabLeft = .Add(msoControlButton)
        tabLeft.Caption = "Set tab order by left to right control order"
        Clipboard.SetData LoadResPicture(105, 0)
        tabLeft.PasteFace
        tabLeft.Enabled = False
        Set TabLeftHandler = VBInstance.Events.CommandBarEvents(tabLeft)
        
        'add the order by controls top position
        Set tabTop = .Add(msoControlButton)
        tabTop.Caption = "Set tab order by top to bottom control order"
        Clipboard.SetData LoadResPicture(106, 0)
        tabTop.PasteFace
        tabTop.Enabled = False
        Set TabTopHandler = VBInstance.Events.CommandBarEvents(tabTop)
        
        'add the plus tab order commandbar
        Set tabMenuPlus = .Add(msoControlButton)
        tabMenuPlus.Caption = "Increase tab index on selected item"
        Clipboard.SetData LoadResPicture(102, 0)
        tabMenuPlus.PasteFace
        tabMenuPlus.Enabled = False
        Set TabPlusHandler = VBInstance.Events.CommandBarEvents(tabMenuPlus)
        
        'add the minus tab order commandbar
        Set tabMenuMinus = .Add(msoControlButton)
        tabMenuMinus.Caption = "Decrease tab index on selected item"
        Clipboard.SetData LoadResPicture(103, 0)
        tabMenuMinus.PasteFace
        tabMenuMinus.Enabled = False
        Set TabMinusHandler = VBInstance.Events.CommandBarEvents(tabMenuMinus)
        
        'add the set tab order commandbar
        Set tabMenuSet = .Add(msoControlButton)
        tabMenuSet.Caption = "Set the tab index order"
        Clipboard.SetData LoadResPicture(104, 0)
        tabMenuSet.PasteFace
        tabMenuSet.Enabled = False
        Set TabSetHandler = VBInstance.Events.CommandBarEvents(tabMenuSet)
        
        'add the container control commandbar
        'Set tabCon = .Add(msoControlButton)
        'tabCon.Caption = "Set tab order on selected container control."
        'Clipboard.SetData LoadResPicture(109, 0)
        'tabCon.PasteFace
        'tabCon.Enabled = False
        'Set ConHandler = VBInstance.Events.CommandBarEvents(tabCon)
        
        'add the set tab order commandbar
        Set tabHelp = .Add(msoControlButton)
        tabHelp.Caption = "Help"
        Clipboard.SetData LoadResPicture(108, 0)
        tabHelp.PasteFace
        tabHelp.Enabled = True
        Set HelpHandler = VBInstance.Events.CommandBarEvents(tabHelp)
        
    End With
    
    With VBInstance.CommandBars("Tab Order")
        .Position = 1
        .Top = 3
        .Left = 3
        .RowIndex = 2
    End With
    
    tabOrderToolbar.Visible = True
    Exit Sub
    
AddToAddInCommandBarErr:
    
    MsgBox "Unable to create toolbar."
    
End Sub

'get all the tab index'(s) and disply them
Private Sub Update()
    Dim i As Integer
    Dim ctl As VBControl
    Dim Count As Integer
    Dim conCtl As VBControl
    Dim subCtl As VBControl
    Dim ti As Integer
    Dim c As Integer
    Dim scaleMode As Integer
    Dim subI As Integer
    Dim subDone As Boolean
    Dim mainC As Integer
    Dim mainCtl As VBControl
    Dim newCnt As Integer
    Dim conLeft As Integer
    Dim conTop As Integer
    
    If VBInstance.ActiveVBProject Is Nothing Then Exit Sub
    
    'load the component
    Set mcmpCurrentForm = VBInstance.SelectedVBComponent
    
    'check to see if we have a valid component
    If mcmpCurrentForm Is Nothing Then
        MsgBox "Select a form first."
      Exit Sub
    End If
    
    'make sure the active component is a form, user control or property page
    If (mcmpCurrentForm.Type <> vbext_ct_VBForm) And _
       (mcmpCurrentForm.Type <> vbext_ct_UserControl) And _
       (mcmpCurrentForm.Type <> vbext_ct_DocObject) And _
       (mcmpCurrentForm.Type <> vbext_ct_PropPage) Then
        MsgBox "Select a form first."
        Exit Sub
    End If
    tabSet = True
    Set controlSel = VBInstance.Events.SelectedVBControlsEvents(VBInstance.ActiveVBProject, mcmpCurrentForm.Designer)
    'display it in the designer
    SendKeys "+{F7}", True
    VBInstance.SelectedVBComponent.DesignerWindow.WindowState = vbext_ws_Maximize
    VBInstance.SelectedVBComponent.DesignerWindow.SetFocus
    
    'get the number of controls on selected form
    Count = mcmpCurrentForm.Designer.VBControls.Count
    scaleMode = mcmpCurrentForm.Properties!scaleMode
    subDone = True
    For i = 1 To Count
        Set ctl = mcmpCurrentForm.Designer.VBControls(i)
        'try to get the tabindex
        On Error Resume Next
        ti = ctl.Properties!tabIndex
        If Err Then
            'doesn't have a tabindex
            Err.Clear
            GoTo SkipIt
        End If
        ReDim Preserve txtIndex(ti)
        On Error Resume Next
        'check if container
        c = ctl.ContainedVBControls.Count
        'if c > 0 then this is a container we need
        'to add our controls to it.
        If c > 0 Then
            Set conCtl = ctl
        ElseIf ctl.Properties!TabStop = False Then
            Err.Clear
            GoTo SkipIt
        End If
        If Err Then
            Err.Clear
        End If
        If subDone Then
            'add to form here
            AddFormControl ti
        Else
            If subCtl Is Nothing Then
                AddContainerControl conCtl, ctl, ti
                conLeft = conCtl.Properties!Left
                conTop = conCtl.Properties!Top
            Else
               AddContainerControl subCtl, ctl, ti
            End If
        End If
        
        With txtIndex(ti)
            .Properties!Left = ctl.Properties!Left
            .Properties!Top = ctl.Properties!Top
            .Properties!ToolTipText = conLeft & "," & conTop
            .Properties!Name = "txtIndex"
            .Properties!Appearance = 0
            Select Case scaleMode
                Case 0, 1
                    .Properties!Height = 250
                    .Properties!Width = 400
                Case 2
                    .Properties!Height = 15
                    .Properties!Width = 19
                Case 3
                    .Properties!Height = 20
                    .Properties!Width = 25
                Case 4
                    .Properties!Height = 1.1
                    .Properties!Width = 3.1
                Case 5
                    .Properties!Height = 0.198
                    .Properties!Width = 0.26
                Case 6
                    .Properties!Height = 5.02
                    .Properties!Width = 6.615
                Case 7
                    .Properties!Height = 0.5
                    .Properties!Width = 0.661
            End Select
            .Properties!ForeColor = &H80
            .Properties!Caption = ctl.Properties!tabIndex
            .Properties!tag = ctl.Properties!tabIndex
        
            newCnt = ctl.Properties!tabIndex
            'put at end of index
            newCnt = newCnt + 2000
            .Properties!tabIndex = newCnt
        End With
        If c > 0 Then
            If mainCtl Is Nothing Then
                mainC = c
                Set mainCtl = ctl
                subI = c
            Else
                subI = c
                Set subCtl = ctl
                subDone = False
            End If
            subDone = False
        ElseIf subI > 0 Then
           subI = subI - 1
        End If
        If subI = 0 Then
            subDone = True
            If mainC > 0 Then
                mainC = mainC - 1
            End If
            Set subCtl = mainCtl
        End If
        
SkipIt:
    Next i

    Err.Clear
    If Not TabAdded Then
        AddTabChange
    End If
    Set ctl = Nothing
    Set conCtl = Nothing
    Set subCtl = Nothing
    Set mainCtl = Nothing
    
    Exit Sub
    
nexter:
    
    MsgBox "Error: " & Err.Description & ", " & Err.Source
    Err.Clear
    If Not TabAdded Then
        AddTabChange
    End If
    
End Sub

Private Sub cmdRemoveControl(svbc As String, ci As Integer)
    Dim c As VBComponent
    Dim p As VBProject
    Dim vbc As VBControl
    Dim sc As String
    Dim sp As String

    On Error Resume Next
    sp = VBInstance.ActiveVBProject
    sc = mcmpCurrentForm.Name
    If sp = "" Or sc = "" Or svbc = "" Then Exit Sub
    Set p = VBInstance.VBProjects.Item(sp)
    Set c = p.VBComponents.Item(sc)
    Set vbf = c.Designer
    Set vbc = vbf.VBControls.Item(ci)
    vbf.VBControls.Remove vbc
     
    tabSet = False
    If Err Then
        Err.Clear
    End If
    
End Sub

Private Sub AddTabChange()
    
    tabMenuPlus.Enabled = True
    tabMenuMinus.Enabled = True
    tabMenuSet.Enabled = True
    tabLeft.Enabled = True
    tabTop.Enabled = True
    tabSelect.Enabled = True
    TabAdded = True
    
End Sub

Private Sub TabMinusHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim ctl As VBControl
    Dim cnt As Integer
    Dim Count As Integer
    Dim index As Integer
    Dim newTab As Integer
    Dim i As Integer
    Dim found As Boolean
    
    Count = mcmpCurrentForm.Designer.VBControls.Count
    For cnt = Count To 1 Step -1
        Set ctl = mcmpCurrentForm.Designer.VBControls.Item(cnt)
        'For Each ctl In mcmpCurrentForm.Designer.VBControls
        If ctl.InSelection Then
            'make sure its our CommandButton
            If ctl.Properties!Name = "txtIndex" Then
                i = ctl.Properties!Caption
                'make sure its greater than 0 and less than totcount
                If i - 1 < 0 Then
                    Beep
                    Exit For
                End If
                'does it have a tab index
                On Error Resume Next
                If Err Then
                    Err.Clear
                Else
                    newTab = ctl.Properties!Caption - 1
                    ctl.Properties!Caption = newTab
                    index = ctl.Properties!tag
                    Set ctl = mcmpCurrentForm.Designer.VBControls.Item(index)
                    ctl.Properties!tabIndex = newTab
                    found = True
                End If
            End If
        End If
    Next cnt
    
    Set ctl = Nothing
    'if user didn't select a control popup msg and exit
    If Not found Then
        Beep
        'MsgBox "Please select a control first."
        Exit Sub
    End If
    Set ctl = Nothing
    Refresh
    
End Sub

Private Sub TabPlusHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim ctl As VBControl
    Dim cnt As Integer
    Dim Count As Integer
    Dim index As Integer
    Dim newTab As Integer
    Dim i As Integer
    Dim found As Boolean
    
    Count = mcmpCurrentForm.Designer.VBControls.Count
    For cnt = Count To 1 Step -1
        Set ctl = mcmpCurrentForm.Designer.VBControls.Item(cnt)
        'For Each ctl In mcmpCurrentForm.Designer.VBControls
        If ctl.InSelection Then
            'make sure its our CommandButton
            If ctl.Properties!Name = "txtIndex" Then
                i = ctl.Properties!Caption
                'make sure its greater than 0 and less than totcount
                If i + 1 > Count Then
                    Beep
                    Exit Sub
                End If
                'does it have a tab index
                On Error Resume Next
                If Err Then
                    Err.Clear
                Else
                    newTab = ctl.Properties!Caption + 1
                    ctl.Properties!Caption = newTab
                    index = ctl.Properties!tag
                    Set ctl = mcmpCurrentForm.Designer.VBControls.Item(index)
                    ctl.Properties!tabIndex = newTab
                    found = True
                End If
            End If
        End If
    Next cnt
    
    Set ctl = Nothing
    'if user didn't select a control popup msg and exit
    If Not found Then
        Beep
        'MsgBox "Please select a control first."
        Exit Sub
    End If
    
    Set ctl = Nothing
    Refresh
    
End Sub

Private Sub TabSetHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim ctl As VBControl
    Dim i As Integer
    Dim Count As Integer
    
    On Error Resume Next
    'unload our CommandButtones
    If tabSet Then
        Count = mcmpCurrentForm.Designer.VBControls.Count
        For i = Count To 1 Step -1
            Set ctl = mcmpCurrentForm.Designer.VBControls.Item(i)
            If ctl.Properties!Name = "txtIndex" Then
                cmdRemoveControl "txtIndex", i
            Else
                ctl.Properties!Locked = True
            End If
            If Err Then
                Err.Clear
            End If
        Next i
    End If
    
    Set ctl = Nothing
    'turn off tabset
    tabSet = False
    
    tabMenuPlus.Enabled = False
    tabMenuMinus.Enabled = False
    tabMenuSet.Enabled = False
    tabLeft.Enabled = False
    tabTop.Enabled = False
    tabSelect.Enabled = False
    TabAdded = False
    If Err Then
        Err.Clear
    End If
    Set ctl = Nothing
    
End Sub

Private Function Refresh() As Integer
    Dim ctl As VBControl
    Dim Count As Integer
    Dim index As Integer
    Dim tabIndex As Integer
    
    Count = mcmpCurrentForm.Designer.VBControls.Count
    For cnt = Count To 1 Step -1
        Set ctl = mcmpCurrentForm.Designer.VBControls.Item(cnt)
        If ctl.Properties!Name = "txtIndex" Then
            'does it have a tab index
            On Error Resume Next
            If Err Then
                Err.Clear
            Else
                index = ctl.Properties!tag
                Set ctl = mcmpCurrentForm.Designer.VBControls.Item(index)
                tabIndex = ctl.Properties!tabIndex
                Set ctl = mcmpCurrentForm.Designer.VBControls.Item(cnt)
                ctl.Properties!Caption = tabIndex
            End If
        End If
    Next cnt
    
    Set ctl = Nothing
    
End Function

Private Sub TabLeftHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim ctl As VBControl
    Dim ctl2 As VBControl
    Dim Count As Integer
    Dim i As Integer
    Dim i2 As Integer
    Dim sendS As String
    Dim sendT As String
    Dim sendL As String
    Dim ti As Integer
    Dim tmpS As String
    Dim conLeft As Integer
    Dim conTop As Integer
    Dim pos As Integer
    Dim tmpStr As String
    
    On Error Resume Next
    Count = mcmpCurrentForm.Designer.VBControls.Count
    Set newCls = Nothing
    Set newCls = New clsSortedCollection
    
    For cnt = Count To Count / 2 Step -1
        Set ctl = mcmpCurrentForm.Designer.VBControls.Item(cnt)
        If ctl.Properties!Name = "txtIndex" Then
            On Error Resume Next
            ti = ctl.Properties!Caption
            If Err Then
                Err.Clear
            Else
            'sort them by top
                sendS = Trim(CStr(ti))
                tmpStr = ctl.Properties!ToolTipText
                pos = InStr(tmpStr, ",")
                conLeft = Val(Mid(tmpStr, 1, pos - 1))
                conTop = Val(Mid(tmpStr, pos + 1, Len(tmpStr) - pos))
                sendT = Trim(CStr(ctl.Properties!Top))
                sendT = sendT + conTop
                sendL = Trim(CStr(ctl.Properties!Left))
                sendL = sendL + conLeft
                newCls.Add sendL, sendT, sendS
                sendS = ""
            End If
        Else
            Exit For
        End If
    Next cnt
    
    'reassign the sort
    For cnt = 1 To newCls.Count
        tmpS = Trim(newCls.Item(2, cnt))
        'put the top here
        sendL = newCls.Item(1, cnt)
        sendT = newCls.Item(3, cnt)
        For i = Count To 1 Step -1
            Set ctl = mcmpCurrentForm.Designer.VBControls.Item(i)
            tmpStr = ctl.Properties!ToolTipText
            pos = InStr(tmpStr, ",")
            conLeft = Val(Mid(tmpStr, 1, pos - 1))
            conTop = Val(Mid(tmpStr, pos + 1, Len(tmpStr) - pos))
            conLeft = conLeft + Val(ctl.Properties!Left)
            conTop = conTop + Val(ctl.Properties!Top)
            If tmpS = ctl.Properties!Caption And sendL = conLeft And sendT = conTop Then
                If Err = 0 Then
                    ctl.Properties!Caption = cnt - 1
                    ti = ctl.Properties!tag
                    'If ti = 0 Then ti = 1
                    ctl.Properties!tag = cnt - 1
                    For i2 = 1 To Count / 2
                        Set ctl2 = mcmpCurrentForm.Designer.VBControls.Item(i2)
                        tmpStr = ctl.Properties!ToolTipText
                        pos = InStr(tmpStr, ",")
                        conLeft = Val(Mid(tmpStr, 1, pos - 1))
                        conTop = Val(Mid(tmpStr, pos + 1, Len(tmpStr) - pos))
                        conLeft = conLeft + Val(ctl2.Properties!Left)
                        conTop = conTop + Val(ctl2.Properties!Top)
                        If sendL = conLeft And sendT = conTop Then
                            ctl2.Properties!tabIndex = cnt - 1
                            Exit For
                        Else
                            Err.Clear
                        End If
                    Next i2
                    Exit For
                Else
                    Err.Clear
                End If
            End If
        Next i
    Next cnt

    If Err Then
        Err.Clear
    End If
    Set newCls = Nothing
    Set ctl = Nothing
    Set ctl2 = Nothing
    
End Sub

Private Sub TabTopHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim ctl As VBControl
    Dim ctl2 As VBControl
    Dim Count As Integer
    Dim i As Integer
    Dim i2 As Integer
    Dim sendS As String
    Dim sendT As String
    Dim sendL As String
    Dim ti As Integer
    Dim tmpS As String
    Dim conLeft As Integer
    Dim conTop As Integer
    Dim pos As Integer
    Dim tmpStr As String
    
    On Error Resume Next
    Count = mcmpCurrentForm.Designer.VBControls.Count
    Set newCls = Nothing
    Set newCls = New clsSortedCollection
    
    For cnt = Count To Count / 2 Step -1
        Set ctl = mcmpCurrentForm.Designer.VBControls.Item(cnt)
        If ctl.Properties!Name = "txtIndex" Then
            On Error Resume Next
            ti = ctl.Properties!Caption
            If Err Then
                Err.Clear
            Else
            'sort them by top
                sendS = Trim(CStr(ti))
                tmpStr = ctl.Properties!ToolTipText
                pos = InStr(tmpStr, ",")
                conLeft = Val(Mid(tmpStr, 1, pos - 1))
                conTop = Val(Mid(tmpStr, pos + 1, Len(tmpStr) - pos))
                sendT = Trim(CStr(ctl.Properties!Top))
                sendT = sendT + conTop
                sendL = Trim(CStr(ctl.Properties!Left))
                sendL = sendL + conLeft
                newCls.Add sendT, sendL, sendS
                sendS = ""
            End If
        Else
            Exit For
        End If
    Next cnt
    
    'reassign the sort
    For cnt = 1 To newCls.Count
        tmpS = Trim(newCls.Item(2, cnt))
        'put the top here
        sendL = newCls.Item(3, cnt)
        sendT = newCls.Item(1, cnt)
        For i = Count To 1 Step -1
            Set ctl = mcmpCurrentForm.Designer.VBControls.Item(i)
            tmpStr = ctl.Properties!ToolTipText
            pos = InStr(tmpStr, ",")
            conLeft = Val(Mid(tmpStr, 1, pos - 1))
            conTop = Val(Mid(tmpStr, pos + 1, Len(tmpStr) - pos))
            conLeft = conLeft + Val(ctl.Properties!Left)
            conTop = conTop + Val(ctl.Properties!Top)
            If tmpS = ctl.Properties!Caption And sendL = conLeft And sendT = conTop Then
                If Err = 0 Then
                    ctl.Properties!Caption = cnt - 1
                    ti = ctl.Properties!tag
                    'If ti = 0 Then ti = 1
                    ctl.Properties!tag = cnt - 1
                    For i2 = 1 To Count / 2
                        Set ctl2 = mcmpCurrentForm.Designer.VBControls.Item(i2)
                        tmpStr = ctl.Properties!ToolTipText
                        pos = InStr(tmpStr, ",")
                        conLeft = Val(Mid(tmpStr, 1, pos - 1))
                        conTop = Val(Mid(tmpStr, pos + 1, Len(tmpStr) - pos))
                        conLeft = conLeft + Val(ctl2.Properties!Left)
                        conTop = conTop + Val(ctl2.Properties!Top)
                        If sendL = conLeft And sendT = conTop Then
                            ctl2.Properties!tabIndex = cnt - 1
                            Exit For
                        Else
                            Err.Clear
                        End If
                    Next i2
                    Exit For
                Else
                    Err.Clear
                End If
            End If
        Next i
    Next cnt

    If Err Then
        Err.Clear
    End If
    Set newCls = Nothing
    Set ctl = Nothing
    Set ctl2 = Nothing
    
End Sub

Private Sub TabSelectHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim Count As Integer
    Dim cnt As Integer
    Dim inSel As Boolean
    Dim ctl As VBControl
    Dim ti As Integer
    Dim i As Integer
    Dim tag As Integer
    Dim ni As Integer
    
    On Error Resume Next
    Count = ctlNew.cnt
    i = 0
    For cnt = 1 To Count
        ni = cnt - 1
        ti = ctlNew.index(cnt)
        Set ctl = mcmpCurrentForm.Designer.VBControls.Item(ti)
        tag = ctl.Properties!tag
        ctl.Properties!Caption = ni
        ctl.Properties!tag = ni
        If tag = 0 Then tag = 1
        Set ctl = mcmpCurrentForm.Designer.VBControls.Item(tag)
        ctl.Properties!tabIndex = cnt - 1
        i = i + 1
        inSel = True
    Next cnt
    
    If Not inSel Then
        Beep
        Exit Sub
    End If
    tabUpdate = True
    RemoveAndUpdate
    If Err Then
        Err.Clear
    End If
    Set ctl = Nothing
    
End Sub

Private Sub AddContainerControl(ByVal ctl As VBControl, ByVal containerControl As VBControl, ByVal ti As Integer)

    Set txtIndex(ti) = ctl.ContainedVBControls.Add("CommandButton", containerControl)
    txtIndex(ti).Properties!Left = 0
    txtIndex(ti).Properties!Top = 0
            
End Sub

Private Sub AddFormControl(ByVal ti As Integer)

    Set txtIndex(ti) = mcmpCurrentForm.Designer.ContainedVBControls.Add("CommandButton", Nothing)

End Sub

Private Sub RemoveAndUpdate()
    Dim i As Integer
    Dim Count As Integer
    Dim ctl As VBControl
    Dim S As String
    
    On Error Resume Next
    i = 0
    If tabSet Then
        Count = mcmpCurrentForm.Designer.VBControls.Count
        For i = Count To 1 Step -1
            Set ctl = mcmpCurrentForm.Designer.VBControls(i)
            If ctl.Properties!Name = "txtIndex" Then
                S = "txtIndex"
                cmdRemoveControl S, i
            End If
        Next i
    End If

    Erase txtIndex
    If tabUpdate Then
        Update
        tabUpdate = False
    End If
    
    If Err Then
        Err.Clear
    End If
    Set ctl = Nothing
    
End Sub
