Imports Microsoft.Office.Interop
Imports System.Threading
Imports System.Text
Imports FinanceApp.HelperClass
Imports FinanceApp.SharedClass
Imports FinanceApp.PublicClass
Imports System.IO
Public Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
Public Delegate Sub FormatReportDelegate(ByRef sender As Object, ByRef e As EventArgs)
Public Class ExportToExcelFile
    Public Property sqlstr As String
    Public Property Directory As String
    Public Property ReportName As String
    Public Property Parent As Object
    Public Property FormatReportCallback As FormatReportDelegate
    Dim myThread As New Threading.Thread(AddressOf DoWork)
    Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable
    Dim AccessFullPath As String
    Dim AccessTableName As String
    Dim SpecificationName As String
    Dim status As Boolean
    Dim Dataset1 As New DataSet
    Public Property Datasheet As Integer = 1
    Public Property mytemplate As String = "\templates\ExcelTemplate.xltx"
    Public Property QueryList As List(Of QueryWorksheet)

    Public Sub New(ByRef Parent As Object, ByRef Sqlstr As String, ByRef Directory As String, ByRef ReportName As String, ByVal FormatReportCallBack As FormatReportDelegate)
        Me.sqlstr = Sqlstr
        Me.Directory = Directory
        Me.ReportName = ReportName
        Me.Parent = Parent
        Me.FormatReportCallback = FormatReportCallBack
    End Sub
    Public Sub New(ByRef Parent As Object, ByRef Sqlstr As String, ByRef Directory As String, ByRef ReportName As String, ByVal FormatReportCallBack As FormatReportDelegate, ByVal PivotCallback As FormatReportDelegate)
        Me.sqlstr = Sqlstr
        Me.Directory = Directory
        Me.ReportName = ReportName
        Me.Parent = Parent
        Me.FormatReportCallback = FormatReportCallBack
        Me.PivotCallback = PivotCallback

    End Sub

    Public Sub New(ByRef Parent As Object, ByRef Sqlstr As String, ByRef Directory As String, ByRef ReportName As String, ByVal FormatReportCallBack As FormatReportDelegate, ByVal PivotCallback As FormatReportDelegate, ByVal AccessFullpath As String, ByVal AccessTableName As String, ByVal SpecificationName As String)
        Me.sqlstr = Sqlstr
        Me.Directory = Directory
        Me.ReportName = ReportName
        Me.Parent = Parent
        Me.FormatReportCallback = FormatReportCallBack
        Me.PivotCallback = PivotCallback
        Me.AccessFullPath = AccessFullpath
        Me.AccessTableName = AccessTableName
        Me.SpecificationName = SpecificationName
    End Sub

    Public Sub New(ByRef Parent As Object, ByRef Sqlstr As String, ByRef Directory As String, ByRef ReportName As String, ByVal FormatReportCallBack As FormatReportDelegate, ByVal PivotCallback As FormatReportDelegate, ByVal datasheet As Integer, ByVal mytemplate As String)
        Me.sqlstr = Sqlstr
        Me.Directory = Directory
        Me.ReportName = ReportName
        Me.Parent = Parent
        Me.FormatReportCallback = FormatReportCallBack
        Me.PivotCallback = PivotCallback
        Me.Datasheet = datasheet
        Me.mytemplate = mytemplate
    End Sub
    Public Sub New(ByRef Parent As Object, ByRef querylist As List(Of QueryWorksheet), ByRef Directory As String, ByRef ReportName As String, ByVal FormatReportCallBack As FormatReportDelegate, ByVal PivotCallback As FormatReportDelegate)
        Me.QueryList = querylist
        Me.Directory = Directory
        Me.ReportName = ReportName
        Me.Parent = Parent
        Me.FormatReportCallback = FormatReportCallBack
        Me.PivotCallback = PivotCallback
    End Sub


    Public Sub Run(ByRef sender As System.Object, ByVal e As System.EventArgs)

        ' FileName = Application.StartupPath & "\PrintOut"
        If Not myThread.IsAlive Then
            Try
                myThread = New System.Threading.Thread(New ThreadStart(AddressOf DoWork))
                myThread.SetApartmentState(ApartmentState.MTA)
                myThread.Start()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Else
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub

    Sub DoWork()
        Dim errMsg As String = String.Empty
        Dim i As Integer = 0
        Dim errSB As New StringBuilder
        Dim sw As New Stopwatch
        sw.Start()
        ProgressReport(2, "Export To Excel..")
        ProgressReport(6, "Marques..")
        status = GenerateReport(Directory, errMsg, Dataset1)
        ProgressReport(5, "Continues..")
        If status Then


            sw.Stop()
            ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            ProgressReport(3, "")

            If MsgBox("File name: " & Directory & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                Process.Start(Directory)
            End If
            ProgressReport(3, "")
            'ProgressReport(4, errSB.ToString)
        Else
            errSB.Append(errMsg) '& vbCrLf)
            ProgressReport(3, errSB.ToString)
        End If
        sw.Stop()


    End Sub

    Private Function GenerateReport(ByRef FileName As String, ByRef errorMsg As String, ByVal dataset1 As DataSet) As Boolean
        Dim myCriteria As String = String.Empty
        Dim result As Boolean = False

        Dim StopWatch As New Stopwatch
        StopWatch.Start()
        'Open Excel
        Application.DoEvents()

        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim hwnd As System.IntPtr
        Try
            'Create Object Excel 
            ProgressReport(2, "CreateObject..")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            'oXl.ScreenUpdating = False
            'oXl.Visible = False
            oXl.DisplayAlerts = False
            ProgressReport(2, "Opening Template...")
            ProgressReport(2, "Generating records..")
            If mytemplate.Contains("172") Then
                oWb = oXl.Workbooks.Open(mytemplate)
            Else
                oWb = oXl.Workbooks.Open(Application.StartupPath & mytemplate)
            End If

            oXl.Visible = False
            'For i = 0 To 6
            '    oWb.Worksheets.Add()
            'Next

            'Dim events As New List(Of ManualResetEvent)()
            'Dim counter As Integer = 0
            ProgressReport(2, "Creating Worksheet...")
            'DATA


            If IsNothing(QueryList) Then
                oWb.Worksheets(Datasheet).select()
                oSheet = oWb.Worksheets(Datasheet)
                ProgressReport(2, "Get records..")

                FillWorksheet(oSheet, sqlstr)
                Dim orange = oSheet.Range("A1")
                Dim lastrow = GetLastRow(oXl, oSheet, orange)


                If lastrow > 1 Then
                    'Delegate for modification
                    'oSheet.Columns("A:A").numberformat = "dd-MMM-yyyy"
                    FormatReportCallback.Invoke(oSheet, New EventArgs)
                End If
            Else
                'Looping from here
                For i = 0 To QueryList.Count - 1
                    Dim myquery = CType(QueryList(i), QueryWorksheet)
                    oWb.Worksheets(myquery.DataSheet).select()
                    oSheet = oWb.Worksheets(myquery.DataSheet)
                    oSheet.Name = myquery.SheetName
                    ProgressReport(2, "Get records..")

                    FillWorksheet(oSheet, myquery.Sqlstr)
                    Dim orange = oSheet.Range("A1")
                    Dim lastrow = GetLastRow(oXl, oSheet, orange)


                    If lastrow > 1 Then
                        'Delegate for modification
                        'oSheet.Columns("A:A").numberformat = "dd-MMM-yyyy"
                        FormatReportCallback.Invoke(oSheet, New EventArgs)
                    End If
                Next


            End If


            PivotCallback.Invoke(oWb, New EventArgs)
            For i = 0 To oWb.Connections.Count - 1
                oWb.Connections(1).Delete()
            Next
            StopWatch.Stop()

            'FileName = FileName & "\" & String.Format("Report" & ReportName & "-{0}-{1}-{2}.xlsx", Today.Year, Format("00", Today.Month), Format("00", Today.Day))
            FileName = FileName & "\" & String.Format(ReportName)
            ProgressReport(3, "")
            ProgressReport(2, "Saving File ..." & FileName)
            'oSheet.Name = ReportName
            If FileName.Contains("xlsm") Then
                oWb.SaveAs(FileName, FileFormat:=Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled)
            Else
                oWb.SaveAs(FileName)
            End If

            ProgressReport(2, "Elapsed Time: " & Format(StopWatch.Elapsed.Minutes, "00") & ":" & Format(StopWatch.Elapsed.Seconds, "00") & "." & StopWatch.Elapsed.Milliseconds.ToString)
            result = True
        Catch ex As Exception
            ProgressReport(3, ex.Message & FileName)
            errorMsg = ex.Message
        Finally
            'oXl.ScreenUpdating = True
            'clear excel from memory
            Try
                oXl.Quit()
                releaseComObject(oSheet)
                releaseComObject(oWb)
                releaseComObject(oXl)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Catch ex As Exception

            End Try

            Try
                'to make sure excel is no longer in memory
                EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try

        End Try
        Return result
    End Function
    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Parent.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Try
                Parent.Invoke(d, New Object() {id, message})
            Catch ex As Exception

            End Try

        Else
            Select Case id
                Case 2
                    Parent.ToolStripStatusLabel1.Text = message
                Case 3
                    Parent.ToolStripStatusLabel2.Text = Trim(message)
                Case 4
                    Parent.close()
                Case 5
                    Parent.ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    Parent.ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
            End Select

        End If

    End Sub

    Public Shared Sub FillWorksheet(ByVal osheet As Excel.Worksheet, ByVal sqlstr As String, Optional ByVal Location As String = "A1")
        'Dim oRange As Excel.Range
        Dim oExCon As String = My.Settings.oExCon ' My.Settings.oExCon.ToString '"ODBC;DSN=PostgreSQL30;"
        oExCon = oExCon.Insert(oExCon.Length, "UID=" & dbadapter1.userid & ";PWD=" & dbadapter1.password)
        Dim oRange As Excel.Range
        oRange = osheet.Range(Location)
        With osheet.QueryTables.Add(oExCon.Replace("Host=", "Server="), oRange)
            'With osheet.QueryTables.Add(oExCon, osheet.Range("A1"))
            .CommandText = sqlstr
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .BackgroundQuery = True
            .RefreshStyle = Excel.XlCellInsertionMode.xlInsertDeleteCells
            .SavePassword = True
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .PreserveColumnInfo = True
            .Refresh(BackgroundQuery:=False)
            Application.DoEvents()
        End With
        oRange = Nothing

        oRange = osheet.Range("1:1")
        oRange = osheet.Range(Location)
        oRange.Select()
        osheet.Application.Selection.autofilter()
        osheet.Cells.EntireColumn.AutoFit()
    End Sub

    Public Shared Function GetLastRow(ByVal oxl As Excel.Application, ByVal osheet As Excel.Worksheet, ByVal range As Excel.Range) As Long
        Dim lastrow As Long = 1
        oxl.ScreenUpdating = False
        Try
            lastrow = osheet.Cells.Find("*", range, , , Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious).Row
        Catch ex As Exception
        End Try
        Return lastrow
        oxl.ScreenUpdating = True
    End Function

    Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)

    End Sub

    Sub FinishReport(ByRef sender As Object, ByRef e As EventArgs)

    End Sub

    Public Shared Sub ExportToExcelAskDirectory(ByRef FileName As String, ByVal Sqlstr As String, Optional ByVal Location As String = "A1", Optional ByVal Template As String = "\templates\ExcelTemplate.xltx", Optional ByVal CreationDate As Boolean = False, Optional ByVal SheetNum As Integer = 1)
        'ask export location
        Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
        DirectoryBrowser.Description = "Which directory do you want to use?"
        If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim myFilename As String = DirectoryBrowser.SelectedPath & "\" & FileName
            'If ExportToExcelFullPath(myFilename, Sqlstr, dbTools, "A4", "\templates\TATemplate", True) Then
            If ExportToExcelFullPath(myFilename, Sqlstr, Location, Template, CreationDate, SheetNum) Then
                If MsgBox("File name: " & myFilename & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                    Process.Start(myFilename)
                End If
            End If
        End If
    End Sub
    Public Shared Function ExportToExcelFullPath(ByRef Filename As String, ByVal Sqlstr As String, Optional ByVal Location As String = "A1", Optional ByVal Template As String = "\templates\ExcelTemplate.xltx", Optional ByVal CreationDate As Boolean = False, Optional ByVal SheetNum As Integer = 1) As Boolean
        Dim result As Boolean = False
        Application.DoEvents()
        Cursor.Current = Cursors.WaitCursor
        Dim source As String = Filename
        Dim StringBuilder1 As New System.Text.StringBuilder
        Dim hwnd As System.IntPtr
        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        'Dim oRange As Excel.Range

        'Need these variable to kill excel
        Dim aprocesses() As Process = Nothing '= Process.GetProcesses
        Dim aprocess As Process = Nothing
        Try
            'Create Object Excel 
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            Application.DoEvents()
            'oXl.Visible = True
            'get process pid
            'aprocesses = Process.GetProcesses
            'For i = 0 To aprocesses.GetUpperBound(0)
            '    If aprocesses(i).MainWindowHandle.ToString = oXl.Hwnd.ToString Then
            '        aprocess = aprocesses(i)
            '        Exit For
            '    End If
            '    Application.DoEvents()
            'Next
            'oXl.Visible = False
            oXl.DisplayAlerts = False
            oWb = oXl.Workbooks.Open(Application.StartupPath & Template)
            'Loop for chart
            oSheet = oWb.Worksheets(SheetNum)
            If CreationDate Then
                oSheet.Cells(1, 1) = "Updated: " & Format(DateTime.Now, "dd MMM yyyy")
            End If

            FillDataSource(oWb, SheetNum, Sqlstr, DbAdapter1, Location)

            Filename = ValidateFileName(System.IO.Path.GetDirectoryName(source), source)
            oWb.SaveAs(Filename)
            result = True

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
            'clear excel from memory
            oXl.Quit()
            'oXl.Visible = True
            releaseComObject(oSheet)
            releaseComObject(oWb)
            releaseComObject(oXl)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            Try
                'If Not aprocess Is Nothing Then
                '    aprocess.Kill()
                'End If
                EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try

            Cursor.Current = Cursors.Default

        End Try
        Return result
    End Function
    Public Shared Function ExportToExcelFullPath(ByRef Filename As String, ByVal Dataset1 As DataSet) As Boolean
        Dim hwnd As System.IntPtr
        Dim result As Boolean = False
        Application.DoEvents()
        Cursor.Current = Cursors.WaitCursor
        Dim source As String = Filename
        Dim StringBuilder1 As New System.Text.StringBuilder

        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim oRange As Excel.Range

        'Need these variable to kill excel
        Dim aprocesses() As Process = Nothing '= Process.GetProcesses
        Dim aprocess As Process = Nothing
        Try
            'Create Object Excel 
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            Application.DoEvents()
            oXl.Visible = True
            'get process pid
            aprocesses = Process.GetProcesses
            For i = 0 To aprocesses.GetUpperBound(0)
                If aprocesses(i).MainWindowHandle.ToString = oXl.Hwnd.ToString Then
                    aprocess = aprocesses(i)
                    Exit For
                End If
                Application.DoEvents()
            Next
            oXl.Visible = False
            oXl.DisplayAlerts = False
            oWb = oXl.Workbooks.Open(Application.StartupPath & "\templates\RawDataTemplate.xltx")
            'Loop for chart
            oSheet = oWb.Worksheets(1)

            'save excel
            If Not ConvertDataForExcel(StringBuilder1, Dataset1.Tables(0).DefaultView) Then
                Return result
            End If
            'put other info
            Dim abc As String = StringBuilder1.ToString

            Clipboard.SetDataObject(StringBuilder1.ToString, False)

            'oRange = oSheet.Range(oSheet.Cells(1, 1), oSheet.Cells(1, 1))
            oRange = oSheet.Range("A1")
            oRange.Select()
            oSheet.Paste()
            oRange = oSheet.Range("1:1")
            'oRange.AutoFilter()
            oSheet.Cells.EntireColumn.AutoFit()
            Filename = ValidateFileName(System.IO.Path.GetDirectoryName(source), source)
            oWb.SaveAs(Filename)
            result = True
            'FormMenu.setBubbleMessage("Export To Excel", "Done")
        Catch ex As Exception
            'MsgBox(ex.Message)

        Finally
            'clear excel from memory
            oXl.Quit()
            'oXl.Visible = True
            releaseComObject(oSheet)
            releaseComObject(oWb)
            releaseComObject(oXl)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            Try
                'If Not aprocess Is Nothing Then
                '    aprocess.Kill()
                'End If
                EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try

            Cursor.Current = Cursors.Default

        End Try
        Return result
    End Function
    Public Shared Function ConvertDataForExcel(ByRef StringBuilderData As System.Text.StringBuilder, ByVal DataView As DataView) As Boolean
        Dim myReturn As Boolean = False
        Dim DataTable As DataTable = DataView.ToTable
        Try
            'Add header
            For i = 0 To DataTable.Columns.Count - 1
                StringBuilderData.Append(DataTable.Columns(i).ToString)
                StringBuilderData.Append(vbTab)
            Next

            StringBuilderData.Append(vbCrLf)
            'Add Detail
            For Each dr In DataTable.Rows

                For i As Long = 0 To dr.itemarray.length - 1
                    '
                    ' Convert the data and fill the string. Null values become blanks.
                    '
                    If dr.itemarray(i).ToString Is DBNull.Value Then
                        StringBuilderData.Append("")
                    Else
                        StringBuilderData.Append(dr.itemarray(i).ToString)
                    End If
                    StringBuilderData.Append(vbTab)
                Next
                '
                ' Add a line feed to the end of each row.
                '
                StringBuilderData.Append(vbCrLf)
                Application.DoEvents()
            Next
            myReturn = True
        Catch ex As Exception
            ' Display an error message.
        End Try
        Return myReturn
    End Function
End Class
Public Class QueryWorksheet
    Public Property DataSheet As Integer
    Public Property Sqlstr As String
    Public Property SheetName As String
End Class
