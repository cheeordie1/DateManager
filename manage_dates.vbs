Option Explicit

Dim xlBookName, outputBookName
Dim xlApp, xlBook, xlOutBook

Set xlApp = CreateObject("Excel.Application")

' Place the full pathname to the excel .xlsx file that
' you want to manage the dates for in the quotes after xlBookName
'
' example on my computer - "C:\Users\Nicholas\Documents\GitHub\DateManager\Calibration_TMA_SOP Schedule_Master Form 101359.xlsx"
'
' (this requires a .xlsx file)

xlBookName = "C:\Users\Nicholas\Documents\GitHub\DateManager\Calibration_TMA_SOP Schedule_Master Form 101359.xlsx"
outputBookName = Replace(xlBookName, ".xlsx", " TODO.xlsx")
Set xlBook = xlApp.Workbooks.Open(xlBookName)

Dim ws, fso

' Check if the TODO file exists, create a new one if not
Set fso = CreateObject("Scripting.FileSystemObject")
If NOT (fso.FileExists(outputBookName)) Then
    Call makeNewTodoBook(xlApp, xlBook, outputBookName)
End If
' Add upcoming deadlines to TODO list
Set xlOutBook = xlApp.Workbooks.Open(outputBookName)
Call updateTodoBook(xlApp, xlBook, xlOutBook)
xlApp.Quit()

' Code to make a new TODO list
Sub makeNewTodoBook(xlApp, xlBook, outputName)
    Dim newBook, newSheet, category, categoryRow
    Set newBook = xlApp.Workbooks.Add()
    For Each ws In xlBook.Worksheets
        Set newSheet = newBook.Worksheets.Add()
        newSheet.Name = ws.Name
        Set categoryRow = ws.Rows(2)
        For category = 1 to categoryRow.Cells.Count
            If (categoryRow.Cells(1,category).Text = "") Then
                Exit For
            End If
            newSheet.Cells(2, category).Value = categoryRow.Cells(1, category).Text
        Next
    Next
    newBook.SaveAs(outputName)
    newBook.Close()
End Sub

' Code to update an old TODO list with new tasks
' If a row entry has a New Calibrate Date that is Within 7 Days
' of the current Date, add it to the TODO list as long as it is
' not currently in the list
Sub updateTodoBook(xlApp, xlBook, xlOutBook)
    Dim curInSheet, curOutSheet
    For Each curInSheet In xlBook.Worksheets
        Set curOutSheet = xlOutBook.Worksheets(curInSheet.Name)
        ' Add Worksheet if it is not already in the output TODO list
        If (curOutSheet Is Nothing) Then
            Set curOutSheet = xlOutBook.Worksheets.Add()
            curOutSheet.Name = curInSheet.Name
        End If

        ' Reduce TODO rows to the output file
        Call exportDateImpendingRows(curInSheet, curOutSheet)
    Next

    xlBook.save()
    xlOutBook.save()
    xlBook.Close()
    xlOutBook.Close()
End Sub

' Function to export the rows of an input sheet that
' have a Next Calibration Date within the next 7 days
Sub exportDateImpendingRows(inSheet, outSheet)
    Dim inTodoIDCol, outTodoIDCol

    inTodoIdCol = findOrAddCategory(inSheet, "TODO ID")
    outTodoIdCol = findOrAddCategory(outSheet, "TODO ID")

End Sub

Function findOrAddCategory(sheet, category)
    Dim categoryCol
    ' Find TODO ID category
    categoryCol = findCategory(sheet, category)

    ' If no TODO ID column, make one
    If (categoryCol = 0) Then
        categoryCol = addCategory(sheet, category)
    End If
    findOrAddCategory = categoryCol
End Function

' Function to find the category with a given name in worksheet
Function findCategory(sheet, category)
    Dim categoryRow, categoryNum
    Dim hasTODOID
    hasTODOID = false
    Set categoryRow = sheet.Rows(2)
    For categoryNum = 1 to categoryRow.Cells.Count
        If (StrComp(categoryRow.Cells(1, categoryNum).Text, "TODO ID") = 0) Then
            hasTODOID = true
            Exit For
        ElseIf (categoryRow.Cells(1, categoryNum).Text = "") Then
            Exit For
        End If
    Next

    If (hasTODOID = false) THEN
        findCategory = 0
    Else
        findCategory = categoryNum
    End If
End Function

' Function to add the category with a given name in worksheet
Sub addCategory(sheet, category)
    Dim numCols, categoryRow
    set categoryRow = sheet.Rows(2)
    For numCols = 1 to categoryRow.Cells.Count
        If (categoryRow.Cells(1, numCols).Text = "") Then
            Exit For
        End If
    Next
    sheet.Cells(2, numCols) = category
    addCategory = numCols
End Sub
