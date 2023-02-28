VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9948
   ClientLeft      =   2412
   ClientTop       =   1212
   ClientWidth     =   6588
   _ExtentX        =   11621
   _ExtentY        =   17547
   _Version        =   393216
   Description     =   "Chnage the deault font for Forms, UserControls, Etc."
   DisplayName     =   "Set default font"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 98 (ver 6.0)"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public FormDisplayed          As Boolean
Public VBInstance             As VBIDE.VBE
Private mcbMenuCommandBar         As Office.CommandBarControl
Private mfrmMain                 As frmMain
Public WithEvents MenuHandler As CommandBarEvents          'controlador de evento de barra de comandos
Attribute MenuHandler.VB_VarHelpID = -1

Private WithEvents mProjects As VBIDE.VBProjectsEvents
Attribute mProjects.VB_VarHelpID = -1
Private WithEvents mComponents As VBIDE.VBComponentsEvents
Attribute mComponents.VB_VarHelpID = -1
Private mNoProject As Boolean

Sub HidefrmMain()
    
    On Error Resume Next
    
    FormDisplayed = False
    If Not mfrmMain Is Nothing Then
        mfrmMain.Hide
    End If
    
End Sub

Sub ShowfrmMain()
  
    On Error Resume Next
    
    If mfrmMain Is Nothing Then
        Set mfrmMain = New frmMain
    End If
    
    Set mfrmMain.VBInstance = VBInstance
    Set mfrmMain.Connect = Me
    FormDisplayed = True
    mfrmMain.Show
    mfrmMain.ZOrder
    mfrmMain.SetFocus
   
End Sub

Private Sub AddinInstance_OnBeginShutdown(custom() As Variant)
    Set mComponents = Nothing
    Set mProjects = Nothing
End Sub

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    Set VBInstance = Application
    
    If ConnectMode = ext_cm_External Then
        ShowfrmMain
    Else
        Set mcbMenuCommandBar = AddToAddInCommandBar(App.Title & " - Configuration")
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    End If
    
    Dim iStr As String
    
    If FontExists("Segoe UI") Then
        iStr = "Segoe UI"
    ElseIf FontExists("Tahoma") Then
        iStr = "Tahoma"
    Else
        iStr = "MS Sans Serif"
    End If
    gNewDefaultFontName = GetSetting(App.Title, "Settings", "DefaultFont", iStr)
    gNewDefaultFontSize = GetSetting(App.Title, "Settings", "DefaultFontSize", "9")
    gChangeDefaultFontEnabled = Val(GetSetting(App.Title, "Settings", "ChangeDefaultFont", "1"))
    
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    mcbMenuCommandBar.Delete
    
    If Not mfrmMain Is Nothing Then
        Unload mfrmMain
        Set mfrmMain = Nothing
    End If
    
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        ShowfrmMain
    End If
End Sub

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
    Set mProjects = Nothing
    Set mProjects = VBInstance.Events.VBProjectsEvents
    SetActiveObjectsHandlers
    If Not VBInstance.ActiveVBProject Is Nothing Then
        If gChangeDefaultFontEnabled And (VBInstance.ActiveVBProject.VBComponents.Count = 1) Then
            If VBInstance.ActiveVBProject.VBComponents(1).Name = "Form1" Then
                ChangeComponentFont VBInstance.ActiveVBProject.VBComponents(1)
            End If
        End If
    End If
End Sub

Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    ShowfrmMain
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'objeto de barra de comandos
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        Exit Function
    End If
    
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function

Private Sub SetActiveObjectsHandlers()
    Set mComponents = Nothing
    If Not VBInstance.ActiveVBProject Is Nothing Then
        Set mComponents = VBInstance.Events.VBComponentsEvents(VBInstance.ActiveVBProject)
    End If
End Sub

Private Sub mProjects_ItemActivated(ByVal VBProject As VBIDE.VBProject)
    SetActiveObjectsHandlers
End Sub

Private Sub mProjects_ItemAdded(ByVal VBProject As VBIDE.VBProject)
    mNoProject = False
    SetActiveObjectsHandlers
    If gChangeDefaultFontEnabled Then
        If VBProject.VBComponents.Count = 1 Then
            ChangeComponentFont VBProject.VBComponents(1)
        End If
    End If
End Sub

Private Sub mProjects_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
    If VBInstance.VBProjects.Count = 1 Then
        mNoProject = True
    End If
    SetActiveObjectsHandlers
End Sub

Private Sub mProjects_ItemRenamed(ByVal VBProject As VBIDE.VBProject, ByVal OldName As String)
    SetActiveObjectsHandlers
End Sub

Private Sub mComponents_ItemAdded(ByVal VBComponent As VBIDE.VBComponent)
    If Not mNoProject Then
        If gChangeDefaultFontEnabled And (Not VBComponent Is Nothing) Then
            ChangeComponentFont VBComponent
        End If
    End If
End Sub

