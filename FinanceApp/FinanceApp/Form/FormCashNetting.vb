Imports System.Threading
Imports Microsoft.Office.Interop
Imports FinanceApp.SharedClass
Imports System.Text
Imports FinanceApp.PublicClass
Public Class FormCashNetting

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""

        Dim mymessage As String = String.Empty
        Dim startdate As Date = DateTimePicker1.Value.Date
        Dim enddate As Date = DateTimePicker2.Value.Date
        Dim sqlstr As String = String.Format("select finance.getcashnetting({0}::date,{1}::date) as result;", DateFormatyyyyMMdd(startdate), DateFormatyyyyMMdd(enddate))

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("CashNettingUploadFile{0:yyyyMMdd}.xlsx", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 1 'because hidden

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\ExcelTemplate.xltx")
            myreport.Run(Me, e)
        End If

    End Sub
   

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)

    End Sub
    Private Sub PivotTable()

    End Sub

    'Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
    '    If Me.InvokeRequired Then
    '        Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
    '        Me.Invoke(d, New Object() {id, message})
    '    Else
    '        Try
    '            Select Case id
    '                Case 1
    '                    ToolStripStatusLabel1.Text = message
    '                Case 2
    '                    ToolStripStatusLabel2.Text = message
    '                Case 4
    '                    'runreport(Me, New System.EventArgs)
    '                Case 5
    '                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
    '                Case 6
    '                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
    '            End Select
    '        Catch ex As Exception

    '        End Try
    '    End If

    'End Sub



End Class