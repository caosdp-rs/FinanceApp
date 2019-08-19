Imports System.Threading
Imports FinanceApp.PublicClass
Imports System.Text
Imports FinanceApp.SharedClass
Public Class FormImportSupplierInvoice

    Dim mythread As New Thread(AddressOf doWork)

    Dim openfiledialog1 As New OpenFileDialog
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)


    Dim DS As DataSet
    Dim myrecord() As String
    Dim myList As New List(Of String())
    Dim FolderBrowserDialog1 As New FolderBrowserDialog
    Dim mySelectedPath As String


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Start Thread
        If Not mythread.IsAlive Then
            'Get file

            If openfiledialog1.ShowDialog = DialogResult.OK Then
                mythread = New Thread(AddressOf doWork)
                mythread.Start()
            End If
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub

    Private Sub doWork()
        'If DbAdapter1.getproglock("FinanceApp.FormImportSAPInvoice", HelperClass1.UserInfo.DisplayName, 1) Then
        '    ProgressReport(2, "This Program is being used by other person")
        '    Exit Sub
        'End If

        Dim sw As New Stopwatch
        Dim DS As New DataSet
        Dim myMessage As String = String.Empty
        sw.Start()
        ProgressReport(1, "Prepare Data...")

        Dim mystr As New StringBuilder
        Dim myInsert As New System.Text.StringBuilder

        Try
            Dim myTextFile As String = String.Empty
            'Dim dir As New IO.DirectoryInfo(mySelectedPath)
            'Dim arrFI As IO.FileInfo() = dir.GetFiles("*.txt")
            'Dim objTFParser As FileIO.TextFieldParser
            Dim myrecord() As String
            Dim tcount As Long = 0
            ProgressReport(6, "Set To Marque")
            '
            myTextFile = openfiledialog1.FileName
            ProgressReport(2, String.Format("Read Text File...{0}", myTextFile))
            Using objTFParser = New FileIO.TextFieldParser(myTextFile)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(Chr(9))
                    .HasFieldsEnclosedInQuotes = True
                    Dim count As Long = 0
                    ProgressReport(1, "Read Data")
                    Do Until .EndOfData
                        myrecord = .ReadFields
                        If count > 0 Then
                            myList.Add(myrecord)
                        End If
                        count += 1
                    Loop
                End With
            End Using

            For i = 0 To myList.Count - 1
                Try
                    If myList(i)(0).Length > 0 Then
                        'vendorcode,groupid
                        'If i = 1110 Then
                        '    'Debug.Print("hello")
                        'End If
                        'Dim reference2 As String = String.Empty
                        'Dim myref = myList(i)(3).Split("-")
                        'If myref.Count > 1 Then
                        '    reference2 = Cleandata(myref(0))
                        'Else
                        '    reference2 = Cleandata(myList(i)(3))
                        'End If
                        myInsert.Append(myList(i)(0) & vbTab &
                                        myList(i)(1) & vbTab &                                        
                                        validdate(myList(i)(2)) & vbTab &
                                        myList(i)(3).Replace(",", "") & vbTab &
                                        myList(i)(4) & vbTab &
                                        myList(i)(5) & vbTab &
                                        myList(i)(6) & vbCrLf)
                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                    Debug.Print("hello")
                End Try

            Next





            ' Next
            ProgressReport(2, "")
            'Update record

            If myInsert.Length > 0 Then
                ProgressReport(1, "Start Add New Records")
                mystr.Append(String.Format("delete from finance.vendorinvoice;"))
                Dim ra As Long = 0
                ra = DbAdapter1.ExNonQuery(mystr.ToString)
                Dim sqlstr As String = "copy finance.vendorinvoice(invoicenumber,month,mydate,iws,gl,dc,cpy) from stdin with null as 'Null';"

                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False

                Try
                    ProgressReport(1, "Add Record Please wait!")


                    errmessage = DbAdapter1.copy(sqlstr, myInsert.ToString, myret)
                    If myret Then
                        ProgressReport(1, "Add Records Done.")
                    Else
                        ProgressReport(1, errmessage)
                    End If
                Catch ex As Exception
                    ProgressReport(1, ex.Message)

                End Try
            End If
            ProgressReport(5, "Set Continuous Again")

        Catch ex As Exception
            ProgressReport(1, ex.Message)
        End Try
        'DbAdapter1.getproglock("FinanceApp.FormImportSAPInvoice", HelperClass1.UserInfo.DisplayName, 0)
        sw.Stop()
        ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))

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

                Case 4
                    'init data
                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee

            End Select

        End If

    End Sub

    Private Function Cleandata(ByVal input As String) As String
        If input.Length = 0 Then
            Return input
        End If
        If IsNumeric(input.Substring(0, 1)) Then
            If Not IsNumeric(input.Substring(input.Length - 1)) Then
                input = input.Substring(0, input.Length - 1)
            End If
        Else
            If input.Substring(0, 1) = "B" Then
                input = input.Substring(0, input.Length - 1)
            End If
        End If
        Return input
    End Function

    Private Function validdate(p1 As String) As String
        If p1 = "" Then
            Return "Null"
        Else
            Return String.Format("'{0:yyyy-MM-dd}'", CDate(p1))
        End If
    End Function
End Class