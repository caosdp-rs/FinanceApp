Imports System.Reflection
Imports FinanceApp.PublicClass

Public Class FormMenu
    Private Sub FormMenu_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            HelperClass1 = New HelperClass
            DbAdapter1 = New DbAdapter

            'HelperClass1.UserInfo.IsAdmin = DbAdapter1.IsAdmin(HelperClass1.UserId)
            'HelperClass1.UserInfo.AllowUpdateDocument = DbAdapter1.AllowUpdateDocument(HelperClass1.UserId)
            Try
                loglogin(DbAdapter1.userid)
            Catch ex As Exception

            End Try
            Me.Text = GetMenuDesc()
            Me.Location = New Point(300, 10)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Me.Close()
        End Try

    End Sub
    Private Sub loglogin(ByVal userid As String)
        Dim applicationname As String = "Finance App"
        Dim username As String = Environment.UserDomainName & "\" & Environment.UserName
        Dim computername As String = My.Computer.Name
        Dim time_stamp As DateTime = Now
        DbAdapter1.loglogin(applicationname, userid, username, computername, time_stamp)
    End Sub
    Public Function GetMenuDesc() As String
        'Label1.Text = "Welcome, " & HelperClass1.UserInfo.DisplayName
        Return "App.Version: " & My.Application.Info.Version.ToString & " :: Server: " & DbAdapter1.ConnectionStringDict.Item("HOST") & ", Database: " & DbAdapter1.ConnectionStringDict.Item("DATABASE") & ", Userid: " & HelperClass1.UserId

    End Function
    Private Sub ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ctrl As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        Dim assembly1 As Assembly = Assembly.GetAssembly(GetType(FormMenu))
        Dim frm As Form = CType(assembly1.CreateInstance(assembly1.GetName.Name.ToString & "." & ctrl.Tag.ToString, True), Form)
        Dim inMemory As Boolean = False
        For i = 0 To My.Application.OpenForms.Count - 1
            If My.Application.OpenForms.Item(i).Name = frm.Name Then
                ExecuteForm(My.Application.OpenForms.Item(i))
                inMemory = True
            End If
        Next
        If Not inMemory Then
            ExecuteForm(frm)
        End If
    End Sub

    Private Sub ExecuteForm(ByVal obj As Windows.Forms.Form)
        With obj
            .WindowState = FormWindowState.Normal
            .StartPosition = FormStartPosition.CenterScreen
            .Show()
            .Focus()
        End With
    End Sub

    Private Sub FormMenu_Load1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.FormMenu_Load(Me, New EventArgs)

        AddHandler ImportSalesAnalysisReportToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler CashNettingToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler BPToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ImportExportSalesAnalysisReportToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        'AddHandler UploadDocumentToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        'AddHandler SupplierDocumentRawDataToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        'AddHandler SupplierCategoryToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        'AddHandler PanelStatusToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        'AddHandler VendorStatusToolStripMenuItem1.Click, AddressOf ToolStripMenuItem_Click
        'AddHandler SupplierPanelToolStripMenuItem1.Click, AddressOf ToolStripMenuItem_Click
        'AddHandler MasterStatusToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        'AddHandler DocumentCountToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        'AddHandler ContractToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        'AddHandler DocumentToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        'AddHandler UserToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        'AddHandler GroupToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        'AddHandler SearchDocumentToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        'AddHandler MasterSupplierToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        'AddHandler MasterVendorToolStripMenuItem1.Click, AddressOf ToolStripMenuItem_Click

        'Admin
        'MasterToolStripMenuItem.Visible = HelperClass1.UserInfo.IsAdmin
        'SupplierDocumentRawDataToolStripMenuItem.Visible = HelperClass1.UserInfo.IsAdmin
        'DocumentCountToolStripMenuItem.Visible = HelperClass1.UserInfo.IsAdmin

    End Sub

    Private Sub FormMenu_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not e.CloseReason = CloseReason.ApplicationExitCall Then
            If MessageBox.Show("Are you sure?", "Exit", MessageBoxButtons.OKCancel) = DialogResult.OK Then
                Me.CloseOpenForm()
                HelperClass1.fadeout(Me)
                DbAdapter1.Dispose()
                HelperClass1.Dispose()
            Else
                e.Cancel = True
            End If
        End If
    End Sub
    Private Sub CloseOpenForm()
        For i = 1 To (My.Application.OpenForms.Count - 1)
            My.Application.OpenForms.Item(1).Close()
        Next
    End Sub



    Private Sub ImportSalesAnalysisReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportSalesAnalysisReportToolStripMenuItem.Click

    End Sub

    Private Sub CashNettingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CashNettingToolStripMenuItem.Click

    End Sub

    Private Sub BPToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BPToolStripMenuItem.Click

    End Sub

    Private Sub ImportExportSalesAnalysisReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportExportSalesAnalysisReportToolStripMenuItem.Click

    End Sub

    Private Sub ImportAccountingDataToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportAccountingDataToolStripMenuItem.Click
        Dim myform = New FormImportSAPInvoice
        myform.ShowDialog()
    End Sub

    Private Sub ImportToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ImportToolStripMenuItem1.Click
        Dim myform = New FormImportSupplierInvoice
        myform.ShowDialog()
    End Sub

    Private Sub SupplierExceptionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SupplierExceptionToolStripMenuItem.Click
        Dim myreport = New SupplierException
        myreport.RunSupplierException()
    End Sub

    Private Sub SAPExceptionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SAPExceptionToolStripMenuItem.Click
        Dim myreport = New SupplierException
        myreport.RunSAPException()
    End Sub

    Private Sub SupplierMatchedToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SupplierMatchedToolStripMenuItem.Click
        Dim myreport = New SupplierException
        myreport.RunSupplierMatched()
    End Sub

    Private Sub SAPMatchedToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SAPMatchedToolStripMenuItem.Click
        Dim myreport = New SupplierException
        myreport.RunSAPMatched()
    End Sub

    Private Sub SAPRawDataToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SAPRawDataToolStripMenuItem.Click
        Dim myreport = New SupplierException
        myreport.RunSAPRawData()
    End Sub
End Class