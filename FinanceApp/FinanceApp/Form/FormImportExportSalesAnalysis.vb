Imports System.Threading
Imports FinanceApp.PublicClass
Imports System.Text
Imports FinanceApp.SharedClass
Public Class FormImportExportSalesAnalysis

    Dim mythread As New Thread(AddressOf doWork)

    Dim openfiledialog1 As New OpenFileDialog
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Public Property groupid As Long = 0

    Private startdate As Date = Date.Today.Date
    Private enddate As Date = Date.Today.Date
    Dim DS As DataSet
    Dim myrecord() As String
    Dim myList As New List(Of String())
    Dim FolderBrowserDialog1 As New FolderBrowserDialog
    Dim mySelectedPath As String
    Dim MissingBPDict As Dictionary(Of String, String)

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Start Thread
        If Not mythread.IsAlive Then
            'Get file
            'startdate = DateTimePicker1.Value.Date
            'enddate = DateTimePicker2.Value.Date
            'If openfiledialog1.ShowDialog = DialogResult.OK Then
            '    mythread = New Thread(AddressOf doWork)
            '    mythread.Start()
            'End If
            ToolStripStatusLabel1.Text = ""
            ToolStripStatusLabel2.Text = ""
            With FolderBrowserDialog1
                .RootFolder = Environment.SpecialFolder.Desktop
                .SelectedPath = "c:\"
                .Description = "Select the source directory"
               
                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    mySelectedPath = .SelectedPath

                    Try
                        mythread = New Thread(AddressOf doWork)
                        'mythread.SetApartmentState(ApartmentState.MTA)
                        mythread.Start()
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                End If
            End With
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub

    Private Sub doWork()
        If DbAdapter1.getproglock("FormIportSalesAnalysis", HelperClass1.UserInfo.DisplayName, 1) Then
            ProgressReport(2, "This Program is being used by other person")
            Exit Sub
        End If

        Dim sw As New Stopwatch
        Dim DS As New DataSet
        Dim myMessage As String = String.Empty
        sw.Start()
        ProgressReport(1, "Prepare Data...")

        If Not fillDataSet(DS, myMessage) Then
            sw.Stop()
            ProgressReport(1, myMessage)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If

        Dim mystr As New StringBuilder
        Dim myInsert As New System.Text.StringBuilder

        Try
            Dim myTextFile As String = String.Empty
            Dim dir As New IO.DirectoryInfo(mySelectedPath)
            Dim arrFI As IO.FileInfo() = dir.GetFiles("*.txt")
            'Dim objTFParser As FileIO.TextFieldParser
            Dim myrecord() As String
            Dim tcount As Long = 0
            ProgressReport(6, "Set To Marque")
            'Dim myindex As Integer
            MissingBPDict = New Dictionary(Of String, String)
            For Each fi As IO.FileInfo In arrFI
                myTextFile = fi.FullName
                ProgressReport(2, String.Format("Read Text File...{0}", fi.FullName))

                Using objTFParser = New FileIO.TextFieldParser(myTextFile)
                    With objTFParser
                        .TextFieldType = FileIO.FieldType.Delimited
                        .SetDelimiters(Chr(9))
                        .HasFieldsEnclosedInQuotes = True
                        Dim count As Long = 0
                        ProgressReport(1, "Read Data")
                        Dim bpcode As String = String.Empty
                        Dim bpName As String = String.Empty
                        Do Until .EndOfData
                            myrecord = .ReadFields

                            'If count = 0 Then 'Check Header
                            '    If myrecord.Length <> 11 Then
                            '        ProgressReport(5, "Continuous")
                            '        ProgressReport(1, "Incorrect File Format.")
                            '        Exit Sub
                            '    End If
                            'Else  'Text file has header
                            If count > 0 Then
                                Dim postingdate As Date
                                Dim duedate As Date
                                Dim doctype As String
                                Dim document As String
                                Dim curr As String
                                'myrecord(1) = CustomerCode

                                If myrecord(1) <> "" Then
                                    Dim mykey(0)
                                    mykey(0) = myrecord(1)
                                    Dim myresult As DataRow = DS.Tables(0).Rows.Find(mykey)
                                    If Not IsNothing(myresult) Then
                                        bpcode = myresult.Item("bpcode").ToString.Substring(3, 3)
                                        bpName = myrecord(2)
                                    Else
                                        If Not MissingBPDict.ContainsKey(myrecord(1)) Then
                                            MissingBPDict.Add(myrecord(1), myrecord(2))
                                        End If
                                    End If
                                Else
                                    postingdate = ValidDateddmmyy(myrecord(7))
                                    duedate = ValidDateddmmyy(myrecord(8))
                                    'Doctype and document                                    
                                    doctype = myrecord(4)
                                    document = myrecord(5)
                                    curr = myrecord(10)

                                    If IsNumeric(myrecord(0)) Then
                                        'vendorcode,groupid
                                        'myInsert.Append(validstr(doctype) & vbTab &
                                        '                validlong(document) & vbTab &
                                        '                validstr(myrecord(2)) & vbTab &
                                        '                validstr(myrecord(3)) & vbTab &
                                        '                DateFormatyyyyMMddString(CStr(postingdate)) & vbTab &
                                        '                DateFormatyyyyMMddString(CStr(duedate)) & vbTab &
                                        '                validstr(myrecord(6)) & vbTab &
                                        '                validreal(myrecord(7)) & vbTab &
                                        '                validreal(myrecord(8)) & vbTab &
                                        '                validreal(myrecord(9)) & vbTab &
                                        '                validreal(myrecord(10)) & vbTab &
                                        '                validstr(bpcode) & vbCrLf)
                                        'doctype,document,installment,salesemployee,postingdate,duedate,customername,salesamount,appliedamount,grossprofit,grossprofitpct,customercode
                                        myInsert.Append(validstr(doctype) & vbTab &
                                                        validlong(document) & vbTab &
                                                        DateFormatyyyyMMddString(CStr(postingdate)) & vbTab &
                                                        DateFormatyyyyMMddString(CStr(duedate)) & vbTab &
                                                        validstr(bpName) & vbTab &
                                                        validreal(myrecord(11)) & vbTab &
                                                        validlong(bpcode) & vbTab &
                                                        validstr(curr) & vbCrLf)

                                    End If
                                End If
                            End If
                            count += 1
                        Loop
                    End With

                End Using


            Next
            ProgressReport(2, "")
            'Update record
            Dim MissingSB As New StringBuilder

            If MissingBPDict.Count > 0 Then
                For i = 0 To MissingBPDict.Count - 1
                    MissingSB.Append(MissingBPDict.Keys(0) & vbCrLf)
                Next
                MessageBox.Show(String.Format("Missing BP : {0}{1}", vbCrLf, MissingSB.ToString))
                Exit Sub
            End If
            If myInsert.Length > 0 Then
                ProgressReport(1, "Start Add New Records")
                'mystr.Append(String.Format("delete from finance.salestx where postingdate >= {0} and postingdate <= {1};", DateFormatyyyyMMdd(startdate), DateFormatyyyyMMdd(enddate)))
                'mystr.Append(String.Format("select finance.deletesalestx({0},{1});", DateFormatyyyyMMdd(startdate), DateFormatyyyyMMdd(enddate)))
                mystr.Append("delete from finance.salestx;")
                'Dim sqlstr As String = "copy finance.salestx(doctype,document,installment,salesemployee,postingdate,duedate,customername,salesamount,appliedamount,grossprofit,grossprofitpct,customercode) from stdin with null as 'Null';"
                Dim sqlstr As String = "copy finance.salestx(doctype,document,postingdate,duedate,customername,salesamount,customercode,curr) from stdin with null as 'Null';"
                Dim ra As Long = 0
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False

                Try
                    'If RadioButton1.Checked Then
                    ProgressReport(1, "Replace Record Please wait!")
                    ra = DbAdapter1.ExNonQuery(mystr.ToString)
                    'End If
                    ProgressReport(1, "Add Record Please wait!")
                    errmessage = DbAdapter1.copy(sqlstr, myInsert.ToString, myret)
                    If myret Then
                        ProgressReport(1, "Add Records Done.")
                    Else
                        ProgressReport(1, errmessage)
                        ProgressReport(5, "Set Continuous Again")
                        Exit Sub
                    End If



                Catch ex As Exception
                    ProgressReport(1, ex.Message)
                    ProgressReport(5, "Set Continuous Again")
                    Exit Sub
                End Try

            End If
            ProgressReport(5, "Set Continuous Again")

            sw.Stop()
            'ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            'export to excelfile
            exportfile(mySelectedPath & "\Result.xlsx")
        Catch ex As Exception
            ProgressReport(1, ex.Message)
            ProgressReport(5, "Set Continuous Again")
            Exit Sub
        End Try
    End Sub
    Private Sub doWork2()
        If DbAdapter1.getproglock("FormIportSalesAnalysis", HelperClass1.UserInfo.DisplayName, 1) Then
            ProgressReport(2, "This Program is being used by other person")
            Exit Sub
        End If

        Dim sw As New Stopwatch
        Dim DS As New DataSet
        Dim myMessage As String = String.Empty
        sw.Start()
        ProgressReport(1, "Prepare Data...")

        If Not fillDataSet(DS, myMessage) Then
            sw.Stop()
            ProgressReport(1, myMessage)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If

        Dim mystr As New StringBuilder
        Dim myInsert As New System.Text.StringBuilder

        Try
            Dim myTextFile As String = String.Empty
            Dim dir As New IO.DirectoryInfo(mySelectedPath)
            Dim arrFI As IO.FileInfo() = dir.GetFiles("*.txt")
            'Dim objTFParser As FileIO.TextFieldParser
            Dim myrecord() As String
            Dim tcount As Long = 0
            ProgressReport(6, "Set To Marque")
            'Dim myindex As Integer
            MissingBPDict = New Dictionary(Of String, String)
            For Each fi As IO.FileInfo In arrFI
                myTextFile = fi.FullName
                ProgressReport(2, String.Format("Read Text File...{0}", fi.FullName))

                Using objTFParser = New FileIO.TextFieldParser(myTextFile)
                    With objTFParser
                        .TextFieldType = FileIO.FieldType.Delimited
                        .SetDelimiters(Chr(9))
                        .HasFieldsEnclosedInQuotes = True
                        Dim count As Long = 0
                        ProgressReport(1, "Read Data")
                        Dim bpcode As String = String.Empty
                        Dim bpName As String = String.Empty
                        Do Until .EndOfData
                            myrecord = .ReadFields

                            'If count = 0 Then 'Check Header
                            '    If myrecord.Length <> 11 Then
                            '        ProgressReport(5, "Continuous")
                            '        ProgressReport(1, "Incorrect File Format.")
                            '        Exit Sub
                            '    End If
                            'Else  'Text file has header
                            If count > 0 Then
                                Dim postingdate As Date
                                Dim duedate As Date
                                Dim doctype As String
                                Dim document As String

                                If myrecord(1) <> "" Then
                                    Dim mykey(0)
                                    mykey(0) = myrecord(1)
                                    Dim myresult As DataRow = DS.Tables(0).Rows.Find(mykey)
                                    If Not IsNothing(myresult) Then
                                        bpcode = myresult.Item("bpcode").ToString.Substring(3, 3)
                                        bpName = myrecord(2)
                                    Else
                                        If Not MissingBPDict.ContainsKey(myrecord(1)) Then
                                            MissingBPDict.Add(myrecord(1), myrecord(2))
                                        End If
                                    End If
                                Else
                                    postingdate = ValidDateddmmyy(myrecord(6))
                                    duedate = ValidDateddmmyy(myrecord(7))
                                    'Doctype and document                                    
                                    doctype = myrecord(3)
                                    document = myrecord(4)

                                    If IsNumeric(myrecord(0)) Then
                                        'vendorcode,groupid
                                        'myInsert.Append(validstr(doctype) & vbTab &
                                        '                validlong(document) & vbTab &
                                        '                validstr(myrecord(2)) & vbTab &
                                        '                validstr(myrecord(3)) & vbTab &
                                        '                DateFormatyyyyMMddString(CStr(postingdate)) & vbTab &
                                        '                DateFormatyyyyMMddString(CStr(duedate)) & vbTab &
                                        '                validstr(myrecord(6)) & vbTab &
                                        '                validreal(myrecord(7)) & vbTab &
                                        '                validreal(myrecord(8)) & vbTab &
                                        '                validreal(myrecord(9)) & vbTab &
                                        '                validreal(myrecord(10)) & vbTab &
                                        '                validstr(bpcode) & vbCrLf)
                                        'doctype,document,installment,salesemployee,postingdate,duedate,customername,salesamount,appliedamount,grossprofit,grossprofitpct,customercode
                                        myInsert.Append(validstr(doctype) & vbTab &
                                                        validlong(document) & vbTab &
                                                        DateFormatyyyyMMddString(CStr(postingdate)) & vbTab &
                                                        DateFormatyyyyMMddString(CStr(duedate)) & vbTab &
                                                        validstr(bpName) & vbTab &
                                                        validreal(myrecord(9)) & vbTab &
                                                        validlong(bpcode) & vbCrLf)

                                    End If
                                End If
                            End If
                            count += 1
                        Loop
                    End With

                End Using


            Next
            ProgressReport(2, "")
            'Update record
            Dim MissingSB As New StringBuilder

            If MissingBPDict.Count > 0 Then
                For i = 0 To MissingBPDict.Count - 1
                    MissingSB.Append(MissingBPDict.Keys(0) & vbCrLf)
                Next
                MessageBox.Show(String.Format("Missing BP : {0}{1}", vbCrLf, MissingSB.ToString))
                Exit Sub
            End If
            If myInsert.Length > 0 Then
                ProgressReport(1, "Start Add New Records")
                'mystr.Append(String.Format("delete from finance.salestx where postingdate >= {0} and postingdate <= {1};", DateFormatyyyyMMdd(startdate), DateFormatyyyyMMdd(enddate)))
                'mystr.Append(String.Format("select finance.deletesalestx({0},{1});", DateFormatyyyyMMdd(startdate), DateFormatyyyyMMdd(enddate)))
                mystr.Append("delete from finance.salestx;")
                'Dim sqlstr As String = "copy finance.salestx(doctype,document,installment,salesemployee,postingdate,duedate,customername,salesamount,appliedamount,grossprofit,grossprofitpct,customercode) from stdin with null as 'Null';"
                Dim sqlstr As String = "copy finance.salestx(doctype,document,postingdate,duedate,customername,salesamount,customercode) from stdin with null as 'Null';"
                Dim ra As Long = 0
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False

                Try
                    'If RadioButton1.Checked Then
                    ProgressReport(1, "Replace Record Please wait!")
                    ra = DbAdapter1.ExNonQuery(mystr.ToString)
                    'End If
                    ProgressReport(1, "Add Record Please wait!")
                    errmessage = DbAdapter1.copy(sqlstr, myInsert.ToString, myret)
                    If myret Then
                        ProgressReport(1, "Add Records Done.")
                    Else
                        ProgressReport(1, errmessage)
                        ProgressReport(5, "Set Continuous Again")
                        Exit Sub
                    End If



                Catch ex As Exception
                    ProgressReport(1, ex.Message)
                    ProgressReport(5, "Set Continuous Again")
                    Exit Sub
                End Try

            End If
            ProgressReport(5, "Set Continuous Again")

            sw.Stop()
            'ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            'export to excelfile
            exportfile(mySelectedPath & "\Result.xlsx")
        Catch ex As Exception
            ProgressReport(1, ex.Message)
            ProgressReport(5, "Set Continuous Again")
            Exit Sub
        End Try
    End Sub
    Private Sub doWork1()
        If DbAdapter1.getproglock("FormIportSalesAnalysis", HelperClass1.UserInfo.DisplayName, 1) Then
            ProgressReport(2, "This Program is being used by other person")
            Exit Sub
        End If

        Dim sw As New Stopwatch
        Dim DS As New DataSet
        Dim myMessage As String = String.Empty
        sw.Start()
        ProgressReport(1, "Prepare Data...")

        If Not fillDataSet(DS, myMessage) Then
            sw.Stop()
            ProgressReport(1, myMessage)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If

        Dim mystr As New StringBuilder
        Dim myInsert As New System.Text.StringBuilder

        Try
            Dim myTextFile As String = String.Empty
            Dim dir As New IO.DirectoryInfo(mySelectedPath)
            Dim arrFI As IO.FileInfo() = dir.GetFiles("*.txt")
            'Dim objTFParser As FileIO.TextFieldParser
            Dim myrecord() As String
            Dim tcount As Long = 0
            ProgressReport(6, "Set To Marque")
            'Dim myindex As Integer
            MissingBPDict = New Dictionary(Of String, String)
            For Each fi As IO.FileInfo In arrFI
                myTextFile = fi.FullName
                ProgressReport(2, String.Format("Read Text File...{0}", fi.FullName))
                Using objTFParser = New FileIO.TextFieldParser(myTextFile)
                    With objTFParser
                        .TextFieldType = FileIO.FieldType.Delimited
                        .SetDelimiters(Chr(9))
                        .HasFieldsEnclosedInQuotes = True
                        Dim count As Long = 0
                        ProgressReport(1, "Read Data")
                        Do Until .EndOfData
                            myrecord = .ReadFields
                            Dim bpcode As String = String.Empty
                            'If count = 0 Then 'Check Header
                            '    If myrecord.Length <> 11 Then
                            '        ProgressReport(5, "Continuous")
                            '        ProgressReport(1, "Incorrect File Format.")
                            '        Exit Sub
                            '    End If
                            'Else  'Text file has header
                            If count > 0 Then
                                Dim PostingDate As Date = ValidDateddmmyy(myrecord(4))
                                Dim duedate As Date = ValidDateddmmyy(myrecord(5))
                                'If PostingDate >= startdate And PostingDate <= enddate Then

                                'Doctype and document
                                Dim mysplit() = myrecord(1).Split(" ")
                                Dim doctype As String = mysplit(0)
                                Dim document As String = mysplit(1)


                                Dim mykey(0)
                                mykey(0) = myrecord(6)
                                Dim myresult As DataRow = DS.Tables(0).Rows.Find(mykey)
                                If Not IsNothing(myresult) Then
                                    bpcode = myresult.Item("bpcode")
                                Else
                                    If Not MissingBPDict.ContainsKey(myrecord(6)) Then
                                        MissingBPDict.Add(myrecord(6), myrecord(6))
                                    End If
                                End If

                                If IsNumeric(myrecord(0)) Then
                                    'vendorcode,groupid
                                    myInsert.Append(validstr(doctype) & vbTab &
                                                    validlong(document) & vbTab &
                                                    validstr(myrecord(2)) & vbTab &
                                                    validstr(myrecord(3)) & vbTab &
                                                    DateFormatyyyyMMddString(CStr(PostingDate)) & vbTab &
                                                    DateFormatyyyyMMddString(CStr(duedate)) & vbTab &
                                                    validstr(myrecord(6)) & vbTab &
                                                    validreal(myrecord(7)) & vbTab &
                                                    validreal(myrecord(8)) & vbTab &
                                                    validreal(myrecord(9)) & vbTab &
                                                    validreal(myrecord(10)) & vbTab &
                                                    validstr(bpcode) & vbCrLf)

                                End If
                                'End If

                            End If
                            count += 1
                        Loop
                    End With

                End Using


            Next
            ProgressReport(2, "")
            'Update record
            Dim MissingSB As New StringBuilder

            If MissingBPDict.Count > 0 Then
                For i = 0 To MissingBPDict.Count - 1
                    MissingSB.Append(MissingBPDict.Keys(0) & vbCrLf)
                Next
                MessageBox.Show(String.Format("Missing BP : {0}{1}", vbCrLf, MissingSB.ToString))
                Exit Sub
            End If
            If myInsert.Length > 0 Then
                ProgressReport(1, "Start Add New Records")
                'mystr.Append(String.Format("delete from finance.salestx where postingdate >= {0} and postingdate <= {1};", DateFormatyyyyMMdd(startdate), DateFormatyyyyMMdd(enddate)))
                'mystr.Append(String.Format("select finance.deletesalestx({0},{1});", DateFormatyyyyMMdd(startdate), DateFormatyyyyMMdd(enddate)))
                mystr.Append("delete from finance.salestx;")
                Dim sqlstr As String = "copy finance.salestx(doctype,document,installment,salesemployee,postingdate,duedate,customername,salesamount,appliedamount,grossprofit,grossprofitpct,customercode) from stdin with null as 'Null';"
                Dim ra As Long = 0
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False

                Try
                    'If RadioButton1.Checked Then
                    ProgressReport(1, "Replace Record Please wait!")
                    ra = DbAdapter1.ExNonQuery(mystr.ToString)
                    'End If
                    ProgressReport(1, "Add Record Please wait!")
                    errmessage = DbAdapter1.copy(sqlstr, myInsert.ToString, myret)
                    If myret Then
                        ProgressReport(1, "Add Records Done.")
                    Else
                        ProgressReport(1, errmessage)
                    End If



                Catch ex As Exception
                    ProgressReport(1, ex.Message)
                    ProgressReport(5, "Set Continuous Again")
                    Exit Sub
                End Try

            End If
            ProgressReport(5, "Set Continuous Again")

            sw.Stop()
            'ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            'export to excelfile
            exportfile(mySelectedPath & "\Result.xlsx")
        Catch ex As Exception
            ProgressReport(1, ex.Message)
            ProgressReport(5, "Set Continuous Again")
            Exit Sub
        End Try

    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Me.ToolStripStatusLabel1.Text = message
                Case 2
                    Me.ToolStripStatusLabel2.Text = message
                Case 3
                    'Report Finished
                    DbAdapter1.getproglock("FormIportSalesAnalysis", HelperClass1.UserInfo.DisplayName, 0)
                Case 4
                    'init data
                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee

            End Select

        End If

    End Sub

    Private Function fillDataSet(ByRef DS As DataSet, ByRef myMessage As String) As Boolean
        'get BP (Business Partner) data
        Dim myret As Boolean = False
        'Dim sqlstr = "select distinct bpname,substr(bpcode::text,4,3) as bpcode from finance.bp order by bpname"
        Dim sqlstr = "select distinct bpname, bpcode from finance.bp order by bpname"

        DS = New DataSet
        If DbAdapter1.TbgetDataSet(sqlstr, DS, myMessage) Then
            Try
                DS.Tables(0).TableName = "BP"
                Dim pk(0) As DataColumn
                'pk(0) = DS.Tables(0).Columns("bpname")
                pk(0) = DS.Tables(0).Columns("bpcode")
                DS.Tables(0).PrimaryKey = pk
            Catch ex As Exception
                myMessage = ex.Message
                Return myret
            End Try
        Else
            Return myret
        End If
        Return True
    End Function

    Private Sub exportfile(myfullpathfilename As String)
        Dim filename = IO.Path.GetDirectoryName(myfullpathfilename)
        Dim reportname = IO.Path.GetFileName(myfullpathfilename)

        Dim datasheet As Integer = 1 'because hidden
        'Dim sqlstr = "select * from finance.salestx;"
        Dim sqlstr As String = String.Format("select finance.getcashnetting({0}::date,{1}::date) as result;", DateFormatyyyyMMdd(CDate("2000-01-31")), DateFormatyyyyMMdd(CDate("9999-12-31")))
        Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
        Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

        Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\ExcelTemplate.xltx")
        myreport.Run(Me, New System.EventArgs)
    End Sub

    Private Sub FormattingReport()
        'Throw New NotImplementedException
    End Sub

    Private Sub PivotTable()
        'Throw New NotImplementedException
    End Sub



End Class