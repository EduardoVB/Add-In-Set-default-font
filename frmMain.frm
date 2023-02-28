VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Config default font"
   ClientHeight    =   2100
   ClientLeft      =   13092
   ClientTop       =   1848
   ClientWidth     =   4908
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4908
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboFontSize 
      Height          =   336
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1032
      Width           =   3252
   End
   Begin VB.CheckBox chkEnabled 
      Caption         =   "Enabled"
      Height          =   396
      Left            =   168
      TabIndex        =   2
      Top             =   1512
      Width           =   2844
   End
   Begin VB.ComboBox cboFont 
      Height          =   336
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   576
      Width           =   3252
   End
   Begin VB.Label Label3 
      Caption         =   "Set the font for new Forms and UserControls:"
      Height          =   324
      Left            =   168
      TabIndex        =   5
      Top             =   168
      Width           =   4548
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Size:"
      Height          =   324
      Left            =   168
      TabIndex        =   3
      Top             =   1032
      Width           =   1164
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Default font:"
      Height          =   324
      Left            =   168
      TabIndex        =   0
      Top             =   624
      Width           =   1164
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public VBInstance As VBIDE.VBE
Public Connect As Connect

Private mLoading As Boolean

Private Sub cboFont_Click()
    If Not mLoading Then
        SaveSetting App.Title, "Settings", "DefaultFont", cboFont.Text
        gNewDefaultFontName = cboFont.Text
    End If
End Sub

Private Sub cboFontSize_Click()
    If Not mLoading Then
        SaveSetting App.Title, "Settings", "DefaultFontSize", cboFontSize.Text
        gNewDefaultFontSize = cboFontSize.Text
    End If
End Sub

Private Sub chkEnabled_Click()
    If Not mLoading Then
        SaveSetting App.Title, "Settings", "ChangeDefaultFont", chkEnabled.Value
        gChangeDefaultFontEnabled = CBool(chkEnabled.Value)
        
        If gChangeDefaultFontEnabled Then
            If Not VBInstance.ActiveVBProject Is Nothing Then
                If gChangeDefaultFontEnabled And (VBInstance.ActiveVBProject.VBComponents.Count = 1) Then
                    ChangeComponentFont VBInstance.ActiveVBProject.VBComponents(1), True
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim c As Long
    Dim iStr As String
    
    mLoading = True
    
    cboFont.Clear
    cboFont.AddItem "MS Sans Serif"
    cboFont.AddItem "Arial"
    cboFont.AddItem "Tahoma"
    cboFont.AddItem "Microsoft Sans Serif"
    If FontExists("Segoe UI") Then cboFont.AddItem "Segoe UI"
    
    For c = 0 To Screen.FontCount - 1
        Select Case Screen.Fonts(c)
            Case "MS Sans Serif", "Arial", "Tahoma", "Microsoft Sans Serif", "Segoe UI"
            Case Else
                cboFont.AddItem Screen.Fonts(c)
        End Select
    Next
    
    If FontExists("Segoe UI") Then
        iStr = "Segoe UI"
    ElseIf FontExists("Tahoma") Then
        iStr = "Tahoma"
    Else
        iStr = "MS Sans Serif"
    End If
    
    SelectInCombo cboFont, gNewDefaultFontName
    
    cboFontSize.Clear
    cboFontSize.AddItem "8"
    cboFontSize.AddItem "9"
    cboFontSize.AddItem "10"
    cboFontSize.AddItem "11"
    cboFontSize.AddItem "12"
    cboFontSize.AddItem "14"
    cboFontSize.AddItem "16"
    cboFontSize.AddItem "18"
    cboFontSize.AddItem "20"
    cboFontSize.AddItem "22"
    cboFontSize.AddItem "24"
    
    SelectInCombo cboFontSize, gNewDefaultFontSize
    
    chkEnabled.Value = Abs(CLng(gChangeDefaultFontEnabled))
    
    Me.Move GetSetting(App.Title, "Settings", "WindowLeft", Screen.Width * 0.7 - Me.Width), GetSetting(App.Title, "Settings", "WindowTop", 1300)
    
    mLoading = False
End Sub

Private Sub SelectInCombo(nCombo As ComboBox, nStr As String)
    Dim c As Long
    
    For c = 0 To nCombo.ListCount - 1
        If nCombo.List(c) = nStr Then
            nCombo.ListIndex = c
            Exit For
        End If
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "Settings", "WindowLeft", Me.Left
    SaveSetting App.Title, "Settings", "WindowTop", Me.Top
End Sub

