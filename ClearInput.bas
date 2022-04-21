Attribute VB_Name = "ClearInput"
Option Explicit

Sub Clear_Input_Data()
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim Answer As VbMsgBoxResult
    
    Answer = MsgBox("Do you want to clear the data on this page?", vbYesNo + vbDefaultButton2, "Clear Data?")
    If Answer = vbYes Then
        Sheets("Input").Activate
        lastRow = Range("A" & Rows.Count).End(xlUp).Row
        lastColumn = Range("A1").End(xlToRight).Column
        Range(Cells(2, 1), Cells(lastRow, lastColumn)).Clear
    End If
    
End Sub

Sub Clear_Output_Data()
    With Sheets("Output")
        .Activate
        .Cells.ClearContents
        .Cells.ClearFormats
    End With
    Sheets("Input").Activate
End Sub
