VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAccessibilityChecklist 
   Caption         =   "Accessibility Checklist"
   ClientHeight    =   9576.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4932
   OleObjectBlob   =   "frmAccessibilityChecklist.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAccessibilityChecklist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Button_DocTitle_Click()
    mSetDocTitle
End Sub

Private Sub Button_AltText_Click()
    mAddAltText
End Sub

Private Sub Button_Headings_Click()
    ' Toggle the Navigation Pane
    If Button_Headings.Value = True Then
        ' Show the Navigation Pane
        Application.CommandBars("Navigation").Visible = True
        Button_Headings.Caption = "Close Navigation Pane"
    Else
        ' Hide the Navigation Pane
        Application.CommandBars("Navigation").Visible = False
        Button_Headings.Caption = "Open Navigation Pane"
    End If
End Sub

Private Sub Button_PBreaks_Click()
    mPBreaks
End Sub

Private Sub Checkbox_DocTitle_Click()

End Sub

Private Sub Checkbox_AltText_Click()

End Sub

Private Sub Checkbox_Headings_Click()

End Sub

Private Sub Checkbox_Links_Click()

End Sub

Private Sub CheckBox_Tables_Click()

End Sub

Private Sub Checkbox_PBreaks_Click()

End Sub

Private Sub Label_DocTitle_Click()

End Sub

Private Sub Label_AltText_Click()

End Sub

Private Sub Label_Headings_Click()

End Sub

Private Sub Label_Links_Click()

End Sub

Private Sub Label_Tables_Click()

End Sub

Private Sub Label_PBreaks_Click()

End Sub

Private Sub Button_Cancel_Click()
    ' Close the form without saving any changes
    Unload Me
End Sub

Private Sub Button_Done_Click()
    On Error Resume Next

    ' Save the checklist status to custom document properties
    ' Alt text
    ActiveDocument.CustomDocumentProperties("AltTextChecked").Value = CheckBox_AltText.Value
    If Err.Number <> 0 Then
        ActiveDocument.CustomDocumentProperties.Add Name:="AltTextChecked", LinkToContent:=False, _
            Type:=msoPropertyTypeBoolean, Value:=CheckBox_AltText.Value
    End If
    Err.Clear

    ' Doc Title
    ActiveDocument.CustomDocumentProperties("DocumentTitleChecked").Value = CheckBox_DocTitle.Value
    If Err.Number <> 0 Then
        ActiveDocument.CustomDocumentProperties.Add Name:="DocumentTitleChecked", LinkToContent:=False, _
            Type:=msoPropertyTypeBoolean, Value:=CheckBox_DocTitle.Value
    End If
    Err.Clear

    ' Paragraph Breaks (PBreaks)
    ActiveDocument.CustomDocumentProperties("NoHardReturnsChecked").Value = CheckBox_PBreaks.Value
    If Err.Number <> 0 Then
        ActiveDocument.CustomDocumentProperties.Add Name:="NoHardReturnsChecked", LinkToContent:=False, _
            Type:=msoPropertyTypeBoolean, Value:=CheckBox_PBreaks.Value
    End If
    Err.Clear
    
    ' Headings
        ActiveDocument.CustomDocumentProperties("UseHeadingsChecked").Value = CheckBox_Headings.Value
    If Err.Number <> 0 Then
        ActiveDocument.CustomDocumentProperties.Add Name:="UseHeadingsChecked", LinkToContent:=False, _
            Type:=msoPropertyTypeBoolean, Value:=CheckBox_Headings.Value
    End If
    Err.Clear
    
    ' Descriptive Links
            ActiveDocument.CustomDocumentProperties("DescriptiveLinksChecked").Value = CheckBox_Links.Value
    If Err.Number <> 0 Then
        ActiveDocument.CustomDocumentProperties.Add Name:="DescriptiveLinksChecked", LinkToContent:=False, _
            Type:=msoPropertyTypeBoolean, Value:=CheckBox_Links.Value
    End If
    Err.Clear
    
    ' Tables
                ActiveDocument.CustomDocumentProperties("TablesChecked").Value = CheckBox_Tables.Value
    If Err.Number <> 0 Then
        ActiveDocument.CustomDocumentProperties.Add Name:="TablesChecked", LinkToContent:=False, _
            Type:=msoPropertyTypeBoolean, Value:=CheckBox_Tables.Value
    End If
    Err.Clear

    ' Close the form
    Unload Me
End Sub



