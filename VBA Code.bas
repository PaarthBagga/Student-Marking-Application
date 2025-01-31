Dim loaded As Boolean
Dim courses() As String
Dim students() As String
Dim grades() As String
Dim selectedFile As Variant

Private Sub btnCancel_Click()
    'Incase cancel button clicked exist
    Unload Me
End Sub

Private Sub btnContinue_Click()
    On Error Resume Next
    'Initialize Variables
    Dim radio As Integer
    Dim conn As Object
    Dim rs As Object
    Dim strSQL As String
    Dim str As String
    Dim counter As Integer
    Dim count As Integer
    Dim code As String
    Dim selectedVal As Variant
    Dim i As Integer
    Dim finalGrade As Double
    Dim min As Double
    Dim max As Double
    Dim average As Double
    Dim mode As Double
    Dim median As Double
    Dim stdDev As Double
    Dim frequencyCount(1 To 10) As Double

    ' Determine which radio button is clicked
    If btnData.Value = True Then
        radio = 1
    ElseIf btnReport.Value = True Then
        radio = 2
    ElseIf btnEnrollment.Value = True Then
        radio = 3
    End If
    'On error go to the handler
    On Error GoTo ErrorHandler
    counter = 0
    
    ' Determine course of action depending on which radio button is clicked
    ' But first ensure that the data is imported in
    If radio = 1 Then
        'Load selected file in through file dialog
        selectedFile = Application.GetOpenFilename("All Files (*.*), *.*", , "Select Database", , False)
        
        ' Get in the file if it's not false
        If selectedFile <> "False" Then
            str = selectedFile
            Set conn = CreateObject("ADODB.Connection")
            conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & str & ";"
            ' Select coursename from courses table
            strSQL = "SELECT CourseName FROM courses;"
            Set rs = CreateObject("ADODB.Recordset")
            rs.Open strSQL, conn
            lbCourses.Clear
            ' Loop through to get values
            Do While Not rs.EOF
                lbCourses.AddItem rs.Fields("CourseName").Value
                count = count + 1
                rs.MoveNext
            Loop
            ReDim Preserve courses(1 To 2, 1 To count)
            count = 0
            ' Create an array to be able to ID each course name
            strSQL = "SELECT CourseCode, CourseName FROM courses;"
            rs.Close
            rs.Open strSQL, conn
            Do While Not rs.EOF
                count = count + 1
                ' Load in the courses array
                courses(1, count) = rs.Fields("CourseCode").Value
                courses(2, count) = rs.Fields("CourseName").Value
                rs.MoveNext
            Loop

            loaded = True
            ' Upon initialization select the first option
            If lbCourses.ListCount > 0 Then
                lbCourses.Selected(0) = True
            End If
            rs.Close
            conn.Close
        ' If cancel clicked send a message
        Else
            MsgBox "Enter a valid database", vbInformation
        End If
        
    ElseIf radio = 2 Or radio = 3 Then
        If loaded = True Then
            If radio = 2 Then
                
                selectedVal = lbCourses.List(lbCourses.ListIndex)
                On Error Resume Next
                Dim wsNew As Worksheet
                Set wsNew = Worksheets(selectedVal)
                On Error GoTo 0
                
                If wsNew Is Nothing Then
                    MsgBox "In order to generate report for this course, first create a worksheet by selecting the 'Course Enrollment' button and clicking continue", vbInformation

                Else
                    Dim doc As Boolean
                    doc = CreateWordReport(wsNew)
                    If Not doc Then
                        Exit Sub
                    End If
                
                End If
            
            
            ElseIf radio = 3 Then
                selectedVal = lbCourses.List(lbCourses.ListIndex)
                ' Error handling if the user tries to create the same sheet again
                On Error Resume Next
                Dim existingSheet As Worksheet
                Set existingSheet = Worksheets(selectedVal)
                On Error GoTo 0
                
                ' If the sheet doesnt exist
                If existingSheet Is Nothing Then
                    Dim newSheet As Worksheet
                    Set newSheet = Worksheets.Add(After:=ThisWorkbook.Sheets("Sheet1"))
                    newSheet.Name = selectedVal
                    
                    ' Find the course code
                    For i = 1 To lbCourses.ListCount
                        If courses(2, i) = selectedVal Then
                            code = courses(1, i)
                        End If
                    Next i
                    
                    ' Access the database to create an array of students
                    str = selectedFile
                    Set conn = CreateObject("ADODB.Connection")
                    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & str & ";"
                    
                    ' Find the number of students
                    strSQL = "SELECT FirstName FROM students;"
                    Set rs = CreateObject("ADODB.Recordset")
                    rs.Open strSQL, conn
                    
                    count = 0
                    ' Loop through to get count
                    Do While Not rs.EOF
                        count = count + 1
                        rs.MoveNext
                    Loop
                    ' First row is Firstname, second is lastname, third is studentId, last is to check if they appear in a course or not
                    ReDim Preserve students(1 To 10, 1 To count)
                    count = 0
                    rs.Close
                    
                    ' Create an array to be able to ID each course name
                    strSQL = "SELECT FirstName, LastName, studentID FROM students;"
                    rs.Open strSQL, conn
                    
                    ' Loop through the database and load in values
                    Do While Not rs.EOF
                        count = count + 1
                        ' Load in the students array
                        students(1, count) = rs.Fields("FirstName").Value
                        students(2, count) = rs.Fields("LastName").Value
                        students(3, count) = rs.Fields("studentID").Value
    
                        rs.MoveNext
                    Loop
                    rs.Close
                    
                    'Also create an array of all grades which includes assignment grades
                    count = 0
                    strSQL = "Select studentID,course,A1,A2,A3,A4,MidTerm,Exam FROM grades;"
                    rs.Open strSQL, conn
                    ' Loop through to get count
                    Do While Not rs.EOF
                        count = count + 1
                        rs.MoveNext
                    Loop
                    
                    'Resize the array appropriately
                    Erase grades
                    ReDim Preserve grades(1 To 8, 1 To count)
                    count = 0
                    rs.Close
                    rs.Open strSQL, conn
                    'Loop in the values
                    Do While Not rs.EOF
                        count = count + 1
                        grades(1, count) = rs.Fields("studentID").Value
                        grades(2, count) = rs.Fields("course").Value
                        grades(3, count) = rs.Fields("A1").Value
                        grades(4, count) = rs.Fields("A2").Value
                        grades(5, count) = rs.Fields("A3").Value
                        grades(6, count) = rs.Fields("A4").Value
                        grades(7, count) = rs.Fields("MidTerm").Value
                        grades(8, count) = rs.Fields("Exam").Value
                        rs.MoveNext
                    Loop
                    
                    
                    rs.Close
                    conn.Close
                    
                    'Set up titles and subtitles on the page
                    newSheet.Cells(1, 1).Value = selectedVal & " Enrollment"
                    newSheet.Cells(2, 1).Value = "First Name: "
                    newSheet.Cells(2, 1).Font.Bold = True
                    
                    newSheet.Cells(2, 2).Value = "Last Name: "
                    newSheet.Cells(2, 2).Font.Bold = True
                    
                    newSheet.Cells(2, 3).Value = "Student ID: "
                    newSheet.Cells(2, 3).Font.Bold = True
                    
                    newSheet.Cells(2, 4).Value = "A1 Grade: "
                    newSheet.Cells(2, 4).Font.Bold = True
                    
                    newSheet.Cells(2, 5).Value = "A2 Grade: "
                    newSheet.Cells(2, 5).Font.Bold = True
                    
                    newSheet.Cells(2, 6).Value = "A3 Grade: "
                    newSheet.Cells(2, 6).Font.Bold = True
                    
                    newSheet.Cells(2, 7).Value = "A4 Grade: "
                    newSheet.Cells(2, 7).Font.Bold = True
                    
                    newSheet.Cells(2, 8).Value = "Midterm Grade: "
                    newSheet.Cells(2, 8).Font.Bold = True
                    
                    newSheet.Cells(2, 9).Value = "Final Exam Grade: "
                    newSheet.Cells(2, 9).Font.Bold = True
                    
                    newSheet.Cells(2, 10).Value = "Final Grade: "
                    newSheet.Cells(2, 10).Font.Bold = True
                    
                    For i = 1 To UBound(students, 2)
                        For j = 1 To UBound(grades, 2)
                            If students(3, i) = grades(1, j) And grades(2, j) = code Then
                                
                                students(4, i) = grades(3, j)
                                students(5, i) = grades(4, j)
                                students(6, i) = grades(5, j)
                                students(7, i) = grades(6, j)
                                students(8, i) = grades(7, j)
                                students(9, i) = grades(8, j)
                                
                                'Calculate the students final grade
                                finalGrade = (Int(grades(3, j)) * 0.05) + (Int(grades(4, j)) * 0.05) + _
                                (Int(grades(5, j)) * 0.05) + (Int(grades(6, j)) * 0.05) + _
                                (Int(grades(7, j)) * 0.3) + (Int(grades(8, j)) * 0.5)
                                
                            
                                students(10, i) = finalGrade
                                
                                'Put the values on the worksheet
                                newSheet.Cells(i + 2, 1).Value = students(1, i)
                                newSheet.Cells(i + 2, 2).Value = students(2, i)
                                newSheet.Cells(i + 2, 3).Value = students(3, i)
                                newSheet.Cells(i + 2, 4).Value = students(4, i)
                                newSheet.Cells(i + 2, 5).Value = students(5, i)
                                newSheet.Cells(i + 2, 6).Value = students(6, i)
                                newSheet.Cells(i + 2, 7).Value = students(7, i)
                                newSheet.Cells(i + 2, 8).Value = students(8, i)
                                newSheet.Cells(i + 2, 9).Value = students(9, i)
                                newSheet.Cells(i + 2, 10).Value = students(10, i)
                                ActiveSheet.UsedRange.Rows.AutoFit
                                ActiveSheet.UsedRange.Columns.AutoFit
                            End If
                        Next j
                    Next i
                        'Collect all statistics needed
                        max = WorksheetFunction.max(newSheet.Range("J3:J52"))
                        min = WorksheetFunction.min(newSheet.Range("J3:J52"))
                        average = WorksheetFunction.average(newSheet.Range("J3:J52"))
                        mode = WorksheetFunction.mode(newSheet.Range("J3:J52"))
                        median = WorksheetFunction.median(newSheet.Range("J3:J52"))
                        stdDev = WorksheetFunction.StDev(newSheet.Range("J3:J52"))
                        
                        
                        For Each cell In newSheet.Range("J3:J52")
                            
                            Value = cell.Value
                            Select Case Value
                                Case 0 To 10
                                    frequencyCount(1) = frequencyCount(1) + 1
                                Case 10.01 To 20
                                    frequencyCount(2) = frequencyCount(2) + 1
                                Case 20.01 To 30
                                    frequencyCount(3) = frequencyCount(3) + 1
                                Case 30.01 To 40
                                    frequencyCount(4) = frequencyCount(4) + 1
                                Case 40.01 To 50
                                    frequencyCount(5) = frequencyCount(5) + 1
                                Case 50.01 To 60
                                    frequencyCount(6) = frequencyCount(6) + 1
                                Case 60.01 To 70
                                    frequencyCount(7) = frequencyCount(7) + 1
                                Case 70.01 To 80
                                    frequencyCount(8) = frequencyCount(8) + 1
                                Case 80.01 To 90
                                    frequencyCount(9) = frequencyCount(9) + 1
                                Case 90.01 To 100
                                    frequencyCount(10) = frequencyCount(10) + 1
                            End Select
                        Next cell
                            newSheet.Cells(1, 11).Value = "Frequency Table:"
                        For i = 1 To 10
                            newSheet.Cells(i, 13).NumberFormat = "0"
                            newSheet.Cells(i, 12).NumberFormat = "@"
                            newSheet.Cells(i, 13).Value = frequencyCount(i)
                            newSheet.Cells(i, 12).Value = CStr((i - 1) * 10) & "-" & CStr(i * 10)
                        Next i
                        
                        
                        Dim HistogramChart As chartObject
                        Set HistogramChart = CreateHistogram(newSheet, code)
                        
                        newSheet.Cells(1, 15).Value = "Max: "
                        newSheet.Cells(2, 15).Value = "Min: "
                        newSheet.Cells(3, 15).Value = "Average: "
                        newSheet.Cells(4, 15).Value = "Mode: "
                        newSheet.Cells(5, 15).Value = "Median: "
                        newSheet.Cells(6, 15).Value = "Std Dev: "
                        newSheet.Cells(1, 16).Value = max
                        newSheet.Cells(2, 16).Value = min
                        newSheet.Cells(3, 16).Value = average
                        newSheet.Cells(4, 16).Value = mode
                        newSheet.Cells(5, 16).Value = median
                        newSheet.Cells(6, 16).Value = stdDev
                Else
                    MsgBox selectedVal & " Enrollment Sheet already exists", vbInformation
                End If
            End If
            
        Else
            MsgBox "You need to import data first", vbCritical
        End If
    End If
ExitProcedure:
    On Error Resume Next
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume ExitProcedure
End Sub
Private Sub UserForm_QueryClose(cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then btnCancel_Click
End Sub

Function CreateHistogram(sheet As Worksheet, code As String) As chartObject
    'Initialize Variables
    Dim histogram As chartObject
    Dim dataRange As Range
    Dim transposedData As Variant
    Dim cell As Range
    Dim HistogramChart As chartObject
    
    Set dataRange = sheet.Range("L1:M10")

    ' Transpose the data range to a variant array
    transposedData = Application.Transpose(dataRange.Value)

    ' Create the histogram chart
    Set histogram = sheet.ChartObjects.Add(Left:=800, Width:=375, Top:=200, Height:=300)
    histogram.Chart.ChartType = xlColumnClustered
    histogram.Chart.SetSourceData Source:=dataRange
    histogram.Chart.HasTitle = True
    histogram.Chart.ChartTitle.Text = "Grade Histogram for " & code
    histogram.Chart.Axes(xlCategory, xlPrimary).HasTitle = True
    histogram.Chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Grades"
    histogram.Chart.Axes(xlValue, xlPrimary).HasTitle = True
    histogram.Chart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Number Of Students"
    histogram.Chart.GapDepth = 0
    histogram.Chart.Legend.Delete
    ' Return the chart
    Set CreateHistogram = histogram

End Function
Private Sub UserForm_Initialize()
    btnData.Value = True
End Sub
Function CreateWordReport(sheet As Worksheet) As Boolean
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim wordChart As Object
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    
    Set wordDoc = wordApp.documents.Add
    wordDoc.Range.Text = sheet.Cells(1, 1).Value & " Report" & vbCrLf & vbCrLf
    
    wordDoc.Range.Text = wordDoc.Range.Text & "Class Statistics are as follows " & vbCrLf
    wordDoc.Range.Text = wordDoc.Range.Text & "Max Final Grade: " & sheet.Cells(1, 16).Value & vbCrLf
    wordDoc.Range.Text = wordDoc.Range.Text & "Min Final Grade: " & sheet.Cells(2, 16).Value & vbCrLf
    wordDoc.Range.Text = wordDoc.Range.Text & "Average Final Grade: " & sheet.Cells(3, 16).Value & vbCrLf
    wordDoc.Range.Text = wordDoc.Range.Text & "Mode of Final Grades: " & sheet.Cells(4, 16).Value & vbCrLf
    wordDoc.Range.Text = wordDoc.Range.Text & "Median of Final Grades: " & sheet.Cells(5, 16).Value & vbCrLf
    wordDoc.Range.Text = wordDoc.Range.Text & "Standard Deviation of Final Grades: " & sheet.Cells(6, 16).Value & vbCrLf
    sheet.ChartObjects(1).CopyPicture
    wordApp.Selection.Paste
    Set wordChart = wordDoc.InlineShapes(1)

    ' Specify the path and filename for saving the Word document
    Dim documentsPath As String
    documentsPath = Environ("USERPROFILE") & "\Documents\"
    
    ' Create a unique filename based on date and time
    Dim fileName As String
    fileName = "Report_" & Format(Now, "YYYYMMDD_HHMMSS") & ".docx"
    
    Set wordDoc = Nothing
    
    
    ' Set the function return value to True (indicating success)
    CreateWordReport = True
    Exit Function

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    ' Set the function return value to False (indicating failure)
    CreateWordReport = False
    Exit Function
End Function
