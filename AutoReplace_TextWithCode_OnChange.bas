' === AutoReplace_TextWithCode_OnChange ===
' Description:
' This macro automatically replaces a value entered by the user
' with a corresponding code from a lookup table (e.g., name ? code).
' Place this code in the worksheet module where the data will be entered.
' * Use Cases:
'   - Converting country names to ISO codes
'   - Replacing job titles with job codes
'   - Translating full text labels into database-friendly IDs
'   - Cleaning pasted data into a standardized format
'--------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim monitoredRange As Range        ' The range to monitor for changes
    Dim changedCell As Range           ' Each changed cell being evaluated
    Dim replacementValue As Variant    ' The code to replace the user's entry

    ' === Customize These Settings ===

    ' 1. Adjust the input range being watched (e.g., columns A to R)
    Set monitoredRange = Intersect(Target, Me.Range("A:R"))
    If monitoredRange Is Nothing Then Exit Sub

    ' 2. Define the lookup range (on sheet "Lists", columns A = Label, B = Code)
    Const lookupSheetName As String = "Lists"
    Const lookupRangeAddress As String = "A:B"

    ' ===============================

    ' Prevent this macro from re-triggering itself
    Application.EnableEvents = False

    ' Loop through all changed cells within the monitored range
    For Each changedCell In monitoredRange
        On Error Resume Next

        ' Try to look up the entered value in the lookup table
        replacementValue = Application.VLookup( _
                              changedCell.Value, _
                              Worksheets(lookupSheetName).Range(lookupRangeAddress), _
                              2, False)

        On Error GoTo 0

        ' If found, replace the cell's value with the code
        If Not IsError(replacementValue) Then
            changedCell.Value = replacementValue
        End If
    Next changedCell

    ' Re-enable event handling
    Application.EnableEvents = True

End Sub




