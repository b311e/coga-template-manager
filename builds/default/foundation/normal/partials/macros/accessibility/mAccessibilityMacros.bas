Attribute VB_Name = "mAccessibilityMacros"
' ===================================================================
' Module: mAccessibilityMacros
' Purpose: Macros for the Accessibility Toolbar
'
' Created 2024-12 by Annabelle Tracy
' Updated 2025-10 by Annabelle Tracy
' ===================================================================
' *******************************************************************

' ===================================================================
' Show Checklist Macro
' ===================================================================

Sub mShowChecklist()
    On Error Resume Next
    Dim prop As Object

    ' Check if the custom property exists, and load its value if it does
    Set prop = ActiveDocument.CustomDocumentProperties("AltTextChecked")
    If Not prop Is Nothing Then
        CheckList.CheckBox_AltText.Value = ActiveDocument.CustomDocumentProperties("AltTextChecked")
    End If

    Set prop = ActiveDocument.CustomDocumentProperties("DocumentTitleChecked")
    If Not prop Is Nothing Then
        CheckList.CheckBox_DocTitle.Value = ActiveDocument.CustomDocumentProperties("DocumentTitleChecked")
    End If

    Set prop = ActiveDocument.CustomDocumentProperties("NoHardReturnsChecked")
    If Not prop Is Nothing Then
        CheckList.CheckBox_PBreaks.Value = ActiveDocument.CustomDocumentProperties("NoHardReturnsChecked")
    End If
    
        Set prop = ActiveDocument.CustomDocumentProperties("UseHeadingsChecked")
    If Not prop Is Nothing Then
        CheckList.CheckBox_Headings.Value = ActiveDocument.CustomDocumentProperties("UseHeadingsChecked")
    End If
    
        Set prop = ActiveDocument.CustomDocumentProperties("DescriptiveLinksChecked")
    If Not prop Is Nothing Then
        CheckList.CheckBox_Links.Value = ActiveDocument.CustomDocumentProperties("DescriptiveLinksChecked")
    End If
    
        Set prop = ActiveDocument.CustomDocumentProperties("TablesChecked")
    If Not prop Is Nothing Then
        CheckList.CheckBox_Tables.Value = ActiveDocument.CustomDocumentProperties("TablesChecked")
    End If

    ' Show the checklist form in modeless mode
    frmAccessibilityChecklist.Show vbModeless
End Sub

' ===================================================================
' Remove Space After Macro
' ===================================================================

Sub mRemoveSpaceAfter()
    ' Ensure that there is text and the cursor is in a paragraph
    If Selection.Type <> wdNoSelection Then
        ' Remove 6 points from the existing space after the current paragraph
        With Selection.ParagraphFormat
            ' Ensure that space after doesn't go below 0
            If .SpaceAfter >= 6 Then
                .SpaceAfter = .SpaceAfter - 6
            Else
                .SpaceAfter = 0
            End If
        End With
    Else
        MsgBox "Please place the cursor in a paragraph to remove spacing.", vbExclamation
    End If
End Sub

' ===================================================================
' Add Space After Macro
' ===================================================================

Sub mAddSpaceAfter()
    ' Ensure that there is text and the cursor is in a paragraph
    If Selection.Type <> wdNoSelection Then
        ' Add 6 points to the existing space after the current paragraph
        With Selection.ParagraphFormat
            .SpaceAfter = .SpaceAfter + 6
        End With
    Else
        MsgBox "Please place the cursor in a paragraph to add spacing.", vbExclamation
    End If
End Sub

' ===================================================================
' Clean Up Paragraph Breaks Macro
' Removes multiple paragraph breaks in a row
' ===================================================================

Sub mPBreaks()
    Dim doc As Document
    Set doc = ActiveDocument

    ' Use a wildcard search to find two or more paragraph marks and replace with one
    With doc.Content.Find
        .ClearFormatting
        .Text = "([!^13])(^13{2,})"
        .replacement.Text = "\1^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    MsgBox "Extra paragraph breaks have been removed.", vbInformation
End Sub

' ===================================================================
' Set Doc Title Macro
' Sets the Document Title metadata field
' ===================================================================

Sub mSetDocTitle()
    Dim doc As Document
    Dim docTitle As String
    Dim currentDocTitle As String

    ' Set reference to the active document
    Set doc = ActiveDocument

    ' Get the current document title
    On Error Resume Next ' Handle cases where the property might not exist
    currentDocTitle = doc.BuiltInDocumentProperties("Title").Value
    On Error GoTo 0

    ' If no title exists, show "(No current title)"
    If currentDocTitle = "" Then
        currentDocTitle = "(No current title)"
    End If

    ' Prompt the user to enter a document title
    docTitle = InputBox("Current Document Title: " & vbCrLf & currentDocTitle & vbCrLf & _
                        "Enter a new document title below:", "Set Document Title", "")

    ' Check if the user entered something
    If docTitle <> "" Then
        ' Set the document title property
        doc.BuiltInDocumentProperties("Title").Value = docTitle
        MsgBox "The document title has been updated to: " & docTitle, vbInformation, "Success"
    Else
        MsgBox "No new title was entered. The document title was not changed.", vbExclamation, "No Changes"
    End If
End Sub

' ===================================================================
' Image Alt Text Macro
' Note: Only works for inline graphics
' ===================================================================

Sub mAddAltText()
    Dim shp As InlineShape
    Dim dlgTitle As String
    Dim currentAltText As String
    Dim newAltText As String
    
    ' Check if a shape is selected
    If Selection.InlineShapes.Count = 1 Then
        Set shp = Selection.InlineShapes(1)
        
        ' Get the current alt text for selected graphic
        currentAltText = shp.AlternativeText
        
        dlgTitle = "Edit Alt Text for Selected Image"
        
        ' Display the input box with an empty field for new alt text
        newAltText = InputBox("Current Alt Text: " & vbCrLf & _
                              IIf(currentAltText = "", "(No current alt text)", currentAltText) & vbCrLf & _
                              "Enter new alt text below:", dlgTitle, "")
        
        ' Update the alt text if user provided a new value
        If newAltText <> "" Then
            shp.AlternativeText = newAltText
            MsgBox "Alt text updated successfully!", vbInformation, "Success"
        Else
            MsgBox "No changes were made.", vbExclamation, "No Changes"
        End If
    Else
        MsgBox "Please select a single image to edit its alt text.", vbExclamation, "No Image Selected"
    End If
End Sub




