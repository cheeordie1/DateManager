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
        Set newSheet = copySheet(ws, newBook)
    Next
    ' Delete the pesky Sheet1
    If (newBook.Worksheets.Count > 1) Then
        xlApp.DisplayAlerts = false
        newBook.Worksheets("Sheet1").Delete
        xlApp.DisplayAlerts = true
    End If
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
        On Error Resume Next
        Set curOutSheet = xlOutBook.Worksheets(curInSheet.Name)
        On Error GoTo 0

        ' Add Worksheet if it is not already in the output TODO list
        If (curOutSheet.Name <> curInSheet.Name) Then
            Set curOutSheet = copySheet(curInSheet, xlOutBook)
        End If

        ' Reduce TODO rows to the output file
        Call exportDateImpendingRows(curInSheet, curOutSheet)
    Next

    xlApp.DisplayAlerts = false
    xlBook.save()
    xlOutBook.save()
    xlApp.DisplayAlerts = true
    xlBook.Close()
    xlOutBook.Close()
End Sub

' Function to copy a worksheet's column types and title into an output sheet
Function copySheet(inSheet, xlOutBook)
    Dim categoryRow, category
    Set copySheet = xlOutBook.Worksheets.Add()
    copySheet.Name = inSheet.Name
    Call copyRow(inSheet, copySheet, 1)
    Call copyRow(inSheet, copySheet, 2)
End Function

' Function to copy a row from one worksheet to another
Sub copyRow(inSheet, outSheet, rowNum)
    inSheet.Rows(rowNum).Copy
    outSheet.Rows(rowNum).PasteSpecial(8)
    outSheet.Rows(rowNum).PasteSpecial(-4163)
    outSheet.Rows(rowNum).PasteSpecial(-4122)
End Sub

' Function to export the rows of an input sheet that
' have a Next Calibration Date within the next 7 days
Sub exportDateImpendingRows(inSheet, outSheet)
    Dim inTodoIDCol, outTodoIDCol
    ' Find the TODO ID column or insert at the end of the sheet
    inTodoIdCol = findOrAddCategory(inSheet, "TODO ID")
    Call synchronizeCategories(inSheet, outSheet)
    outTodoIdCol = findOrAddCategory(outSheet, "TODO ID")

    ' Give TODO IDs to all input rows without them
    Call fillTodoIDs(inSheet, inTodoIdCol)

    ' Reduce all the rows from input sheet with dates within 7 days
    Dim inFutureShortDateCol, inPastShortDateCol, inFutureLongDateCol, inPastLongDateCol

    ' Find the Columns with specific Date Categories to look for
    inFutureShortDateCol = findOneOfCategory(inSheet, Array("Next Calibration Date", "Calibration Due Date", "Next Due"))
    inPastShortDateCol = findOneOfCategory(inSheet, Array("Last Calibration Date", "Date Calibrated"))
    inFutureLongDateCol = findOneOfCategory(inSheet, Array("Next TMA Date"))
    inPastLongDateCol = findOneOfCategory(inSheet, Array("TMA Date"))

    Dim curRow, curRowNum, inNotDataCol, continue, todayDate
    Dim futureLongDate, futureShortDate, reduceRow
    todayDate = Date
    inNotDataCol = findCategory(inSheet, "NOT DATA")
    For curRowNum = 3 to inSheet.UsedRange.Rows.Count
        continue = false
        Set curRow = inSheet.UsedRange.Rows(curRowNum)
        
        ' Stop at first empty row with no NOT DATA value
        If (curRow.Cells.Find("*") is Nothing) Then
            Exit For
        End If

        ' Skip NOT DATA rows
        If (inNotDataCol > 0) Then
            If (NOT (IsEmpty(curRow.Cells(1, inNotDataCol)))) Then
                continue = true
            End If
        End If

        ' Check for if the Row's date column is within 7 days of this date
        If (inFutureShortDateCol <> 0) Then
            futureShortDate = parseDate(curRow.Cells(1, inFutureShortDateCol))
        Else
            futureShortDate = ""
        End If
        If (inFutureLongDateCol <> 0) Then
            futureLongDate = parseDate(curRow.Cells(1, inFutureLongDateCol))
        Else
            futureLongDate = ""
        End If

        reduceRow = false
        If (futureShortDate <> "") Then
            If (Abs(DateDiff("d", todayDate, CDate(futureShortDate))) < 7) Then
                reduceRow = true
            End If
        End If
        If (futureLongDate <> "") Then
            If (Abs(DateDiff("d", todayDate, CDate(futureLongDate))) < 7) Then
                reduceRow = true
            End If
        End If

        ' If reduceRow is true, reduce the row
        If (reduceRow = true) Then
            Call reduceRowToSheet(curRow, outSheet, inTodoIdCol, outTodoIdCol)
        End If

    Next
End Sub

' Add Row to sheet if the output sheet doesn't have the same row already
' Check todoIDs to see if the output sheet contains the row
Sub reduceRowToSheet(row, outSheet, inTodoIdCol, outTodoIdCol)
    ' Check to see if outSheet already contains the TODO ID
    Dim checkID, curCol
    checkID = row.Cells(1,inTodoIdCol).Value
    If (outSheet.UsedRange.Columns(outTodoIdCol).Find(checkID) Is Nothing) Then
        ' If TODO ID is not already present, print the row to the first line of the todo list
        outSheet.Rows(3).Insert(-4121)
        For curCol =  1 to row.Cells.Count
            row.Cells(1,curCol).Copy(outSheet.Rows(3).Cells(1, curCol))
        Next
    End If
End Sub

' Function that parses a date from a cell Returns nothing if there is no valid date
Function parseDate(cell)
    Dim regex, result
    Set regex = new RegExp
    regex.pattern = "(\d{1,2}\/\d{1,2}\/\d{4})|" & _
                    "(\d{1,2}\.\d{1,2}\.\d{4})|" & _ 
                    "(\d{1,2}\-\d{1,2}\-\d{4})|" & _ 
                    "(\d{4})\-\d{1,2}\-\d{1,2}"
    Set result = regex.execute(cell.Text)
    If (regex.test(cell.Text)) Then
        parseDate = result(0).value
    Else
        parseDate = ""
    End if
End Function


' Function that synchronizes the categories from the original xlsx to the
' TODO list. Do this later, because it will take a long time.
Sub synchronizeCategories(inSheet, outSheet)
End Sub

' Function that searches for a given category in a category column
' If not found, adds the category column
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

' Find the column position of one of the categories in a list
Function findOneOfCategory(sheet, categoryArr)
    Dim categoryCol, category, categoryRow
    Set categoryRow = sheet.UsedRange.Rows(2)
    findOneOfCategory = 0
    For Each category in categoryArr
        categoryCol = findCategory(sheet, category)
        If (categoryCol > 0) Then
            findOneOfCategory = categoryCol
            Exit For
        End If
    Next
End Function

' Function to find the category with a given name in worksheet
Function findCategory(sheet, category)
    Dim categoryRow, categoryNum
    Dim hasCategory
    hasCategory = false

    Set categoryRow = sheet.UsedRange.Rows(2)

    For categoryNum = 1 to categoryRow.Cells.Count
        If (StrComp(CStr(categoryRow.Cells(1, categoryNum)), category, vbTextCompare) = 0) Then
            hasCategory = true
            Exit For
        ElseIf (categoryRow.Cells(1, categoryNum).Text = "") Then
            Exit For
        End If
    Next

    If (hasCategory = false) THEN
        findCategory = 0
    Else
        findCategory = categoryNum
    End If
End Function

' Function to add the category with a given name in worksheet
Function addCategory(sheet, category)
    Dim numCols, categoryRow
    set categoryRow = sheet.Rows(2)
    For numCols = 1 to categoryRow.Cells.Count
        If (categoryRow.Cells(1, numCols).Text = "") Then
            Exit For
        End If
    Next
    sheet.Cells(2, numCols) = category
    addCategory = numCols
End Function

' Function to fill a TODO ID column with IDs that will later map to TODO list tasks
Sub fillTodoIDs(sheet, todoColumn)
    Dim nextTodoID, curRowNum, curRow, notDataColumn, continue
    notDataColumn = findCategory(sheet, "NOT DATA")
    nextTodoID = 1
    ' Get the next TODO ID
    For curRowNum = 3 to sheet.UsedRange.Rows.Count
        continue = false
        Set curRow = sheet.UsedRange.Rows(curRowNum)
        
        ' Stop at first empty row with no NOT DATA value
        If (curRow.Cells.Find("*") is Nothing) Then
            Exit For
        End If

        ' Skip NOT DATA rows
        If (notDataColumn > 0) Then
            If (NOT (IsEmpty(curRow.Cells(1, notDataColumn)))) Then
                continue = true
            End If
        End If

        ' If column has a TODO ID, set the max TODO ID respectively
        If (continue = false) Then
            If (NOT (IsEmpty(curRow.Cells(1, todoColumn)))) Then
                If (curRow.Cells(1, todoColumn).Value2 >= nextTodoID) Then
                    nextTodoID = curRow.Cells(1, todoColumn).Value2 + 1
                End If
            End If
        End If
    Next

    ' For each row without a TODO ID, add one
    For curRowNum = 3 to sheet.UsedRange.Rows.Count
        continue = false
        Set curRow = sheet.UsedRange.Rows(curRowNum)

        ' Stop at first empty row with no NOT DATA value
        If (curRow.Cells.Find("*") is Nothing) Then
            Exit For
        End If


        ' Skip NOT DATA rows
        If (notDataColumn > 0) Then
            If (NOT (IsEmpty(curRow.Cells(1, notDataColumn)))) Then
                continue = true
            End If
        End If

        ' If column has a TODO ID, set the max TODO ID respectively
        If (continue = false) Then
            If (IsEmpty(curRow.Cells(1, todoColumn))) Then
                curRow.Cells(1, todoColumn) = nextTodoID
                nextTodoID = nextTodoID + 1
            End If
        End If
    Next
End Sub
