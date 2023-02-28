Attribute VB_Name = "Module1"
Option Explicit

Public gNewDefaultFontName As String
Public gNewDefaultFontSize As String
Public gChangeDefaultFontEnabled As Boolean

Public Function FontExists(nFontName As String) As Boolean
    Dim iFont As New StdFont
    
    iFont.Name = nFontName
    FontExists = StrComp(nFontName, iFont.Name, vbTextCompare) = 0
End Function

Public Sub ChangeComponentFont(nComponent As VBComponent, Optional nReload As Boolean)
    Dim iLng As Long
    
    If nComponent.FileNames(1) = "" Then
        If (nComponent.Type = vbext_ct_VBForm) Or (nComponent.Type = vbext_ct_UserControl) Or (nComponent.Type = vbext_ct_VBMDIForm) Or (nComponent.Type = vbext_ct_PropPage) Then
            iLng = -1
            On Error Resume Next
            iLng = nComponent.Designer.VBControls.Count
            On Error GoTo 0
            If (iLng = 0) Then
                Dim iProp As Property
                Dim p As Long
                Dim iObj As Object
                Const cOrigFontName As String = "MS Sans Serif"
                Const cOrigFontSize As Long = 8
                
                Set iProp = Nothing
                For p = 1 To nComponent.Properties.Count
                    Set iProp = nComponent.Properties(p)
                    If iProp.Name = "Font" Then
                        Set iObj = Nothing
                        On Error Resume Next
                        Set iObj = iProp.object
                        On Error GoTo 0
                        If Not iObj Is Nothing Then
                            If TypeName(iProp.object) = "Font" Then
                                If (iObj.Name = cOrigFontName) And (Round(iObj.Size) = cOrigFontSize) Then
                                    iObj.Name = gNewDefaultFontName
                                    iObj.Size = Val(gNewDefaultFontSize)
                                End If
                            End If
                        End If
                        Set iObj = Nothing
                    End If
                Next
                Set iProp = Nothing 'Bug in the Add-In environment, if not set to Nothing VB chashes with UserControls when the Add-In is compiled
            
                If nComponent.DesignerWindow.Visible And nReload Then
                    nComponent.DesignerWindow.Close
                    nComponent.DesignerWindow.Visible = True
                End If
            
            End If
        End If
    End If
End Sub

