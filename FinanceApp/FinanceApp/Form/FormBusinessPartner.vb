Imports System.Threading
Imports System.Text
Imports FinanceApp.PublicClass
Imports FinanceApp.SharedClass
Imports System.ComponentModel

Public Class FormBusinessPartner
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim WithEvents BPBS As BindingSource

    Dim DS As DataSet
    Dim sb As New StringBuilder

    Dim bsVendorName As BindingSource
    Dim bsVendorNameHelper As BindingSource

    Dim myDict As Dictionary(Of String, Integer)
    Dim myFields As String() = {"bpcode", "bpname"}

    Dim BPTypeList As List(Of BPType)
    Dim CreditTermList As List(Of CreditTerm)
    Dim PriceListCodeList As List(Of PriceListCode)

    Dim BPTypeBS As BindingSource
    Dim CreditTermBS As BindingSource
    Dim PriceListCodeBS As BindingSource

    Private Sub FormSupplierCategory_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        BPTypeList = New List(Of BPType)
        CreditTermList = New List(Of CreditTerm)
        PriceListCodeList = New List(Of PriceListCode)

        BPTypeList.Add(New BPType(DBNull.Value, ""))
        BPTypeList.Add(New BPType("Customer", "Customer"))
        BPTypeList.Add(New BPType("Vendor", "Vendor"))

        CreditTermList.Add(New CreditTerm(DBNull.Value, ""))
        CreditTermList.Add(New CreditTerm("30 DAYS", "30 DAYS"))
        CreditTermList.Add(New CreditTerm("90 DAYS", "90 DAYS"))

        PriceListCodeList.Add(New PriceListCode(DBNull.Value, ""))
        PriceListCodeList.Add(New PriceListCode("P01 - TP init - SGD", "P01 - TP init - SGD"))
        PriceListCodeList.Add(New PriceListCode("P02 - TP real - PurchaseCurrency", "P02 - TP real - PurchaseCurrency"))
        PriceListCodeList.Add(New PriceListCode("PL40 - Export (Malaysia) - SGD", "PL40 - Export (Malaysia) - SGD"))
        loaddata()
    End Sub

    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")

        DS = New DataSet

        Dim mymessage As String = String.Empty
        sb.Clear()
        sb.Append(String.Format("select bpcode::text,bpname,bpbalance,kam,bpbalanceinfc,currentmlacode,creditlimit,creditterm,bptype,pricelistcode,telephone,fax,shiptobuilding,shiptoblock,shiptostreet,shiptozipcode,cofacecoverage,active,bpid from finance.bp order by bpcode;"))

        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try
                DS.Tables(0).TableName = "BP"
                
            Catch ex As Exception
                ProgressReport(1, "Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
            ProgressReport(4, "InitData")
        Else
            ProgressReport(1, "Loading Data. Error::" & mymessage)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If

        ProgressReport(5, "Continuous")
    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else

            Select Case id
                Case 1
                    ToolStripStatusLabel1.Text = message
                Case 2
                    ToolStripStatusLabel1.Text = message
                Case 4


                    BPBS = New BindingSource
                    BPTypeBS = New BindingSource
                    BPTypeBS.DataSource = BPTypeList

                    CreditTermBS = New BindingSource
                    CreditTermBS.DataSource = CreditTermList

                    PriceListCodeBS = New BindingSource
                    PriceListCodeBS.DataSource = PriceListCodeList

                    Dim pk(0) As DataColumn
                    pk(0) = DS.Tables(0).Columns("bpid")
                    DS.Tables(0).PrimaryKey = pk
                    DS.Tables(0).Columns("bpid").AutoIncrement = True
                    DS.Tables(0).Columns("bpid").AutoIncrementSeed = 0
                    DS.Tables(0).Columns("bpid").AutoIncrementStep = -1

                    DS.Tables(0).TableName = "BP"

                    BPBS.DataSource = DS.Tables(0)

                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.DataSource = BPBS
                    DataGridView1.RowTemplate.Height = 22

                    TextBox1.DataBindings.Clear()
                    TextBox2.DataBindings.Clear()
                    TextBox3.DataBindings.Clear()
                    TextBox4.DataBindings.Clear()
                    TextBox5.DataBindings.Clear()
                    TextBox6.DataBindings.Clear()
                    TextBox7.DataBindings.Clear()
                    TextBox8.DataBindings.Clear()
                    TextBox9.DataBindings.Clear()
                    TextBox10.DataBindings.Clear()
                    TextBox11.DataBindings.Clear()


                    TextBox15.DataBindings.Clear()

                    TextBox18.DataBindings.Clear()
                    CheckBox1.DataBindings.Clear()
                    CheckBox2.DataBindings.Clear()
                    ComboBox1.DataBindings.Clear()
                    ComboBox2.DataBindings.Clear()
                    ComboBox3.DataBindings.Clear()

                    TextBox1.DataBindings.Add("text", BPBS, "bpcode", True, DataSourceUpdateMode.OnPropertyChanged)
                    TextBox2.DataBindings.Add("text", BPBS, "bpname", True, DataSourceUpdateMode.OnPropertyChanged)
                    TextBox4.DataBindings.Add("text", BPBS, "kam", True, DataSourceUpdateMode.OnPropertyChanged)
                    TextBox6.DataBindings.Add("text", BPBS, "currentmlacode", True, DataSourceUpdateMode.OnPropertyChanged)

                    TextBox3.DataBindings.Add("text", BPBS, "bpbalance", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00")
                    TextBox5.DataBindings.Add("text", BPBS, "bpbalanceinfc", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00")
                    TextBox18.DataBindings.Add("text", BPBS, "creditlimit", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00")

                    TextBox11.DataBindings.Add("text", BPBS, "telephone", True, DataSourceUpdateMode.OnPropertyChanged)
                    TextBox10.DataBindings.Add("text", BPBS, "fax", True, DataSourceUpdateMode.OnPropertyChanged)
                    TextBox9.DataBindings.Add("text", BPBS, "shiptobuilding", True, DataSourceUpdateMode.OnPropertyChanged)
                    TextBox8.DataBindings.Add("text", BPBS, "shiptoblock", True, DataSourceUpdateMode.OnPropertyChanged)
                    TextBox7.DataBindings.Add("text", BPBS, "shiptostreet", True, DataSourceUpdateMode.OnPropertyChanged)
                    TextBox15.DataBindings.Add("text", BPBS, "shiptozipcode", True, DataSourceUpdateMode.OnPropertyChanged)


                    ComboBox1.DataSource = BPTypeBS
                    ComboBox1.DisplayMember = "BPTypename"
                    ComboBox1.ValueMember = "BPType"
                    ComboBox1.DataBindings.Add("SelectedValue", BPBS, "bptype", True, DataSourceUpdateMode.OnPropertyChanged)

                    ComboBox2.DataSource = CreditTermList
                    ComboBox2.DisplayMember = "CreditTerm"
                    ComboBox2.ValueMember = "CreditTermName"
                    ComboBox2.DataBindings.Add("SelectedValue", BPBS, "creditterm", True, DataSourceUpdateMode.OnPropertyChanged)

                    ComboBox3.DataSource = PriceListCodeList
                    ComboBox3.DisplayMember = "PriceListCode"
                    ComboBox3.ValueMember = "PriceListName"
                    ComboBox3.DataBindings.Add("SelectedValue", BPBS, "pricelistcode", True, DataSourceUpdateMode.OnPropertyChanged)

                    CheckBox1.DataBindings.Add("checked", BPBS, "active", True, DataSourceUpdateMode.OnPropertyChanged)
                    CheckBox2.DataBindings.Add("checked", BPBS, "cofacecoverage", True, DataSourceUpdateMode.OnPropertyChanged)

                    If IsNothing(BPBS.Current) Then
                        ComboBox1.SelectedIndex = -1
                    End If

                   

                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
            End Select
           
        End If

    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        loaddata()
    End Sub

    Private Sub loaddata()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub SCBS_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles BPBS.ListChanged
        TextBox1.Enabled = Not IsNothing(BPBS.Current)
        TextBox2.Enabled = Not IsNothing(BPBS.Current)
        TextBox3.Enabled = Not IsNothing(BPBS.Current)
        TextBox4.Enabled = Not IsNothing(BPBS.Current)
        TextBox5.Enabled = Not IsNothing(BPBS.Current)
        TextBox6.Enabled = Not IsNothing(BPBS.Current)
        TextBox7.Enabled = Not IsNothing(BPBS.Current)
        TextBox8.Enabled = Not IsNothing(BPBS.Current)
        TextBox9.Enabled = Not IsNothing(BPBS.Current)
        TextBox10.Enabled = Not IsNothing(BPBS.Current)
        TextBox11.Enabled = Not IsNothing(BPBS.Current)
        TextBox15.Enabled = Not IsNothing(BPBS.Current)
        TextBox18.Enabled = Not IsNothing(BPBS.Current)
        CheckBox1.Enabled = Not IsNothing(BPBS.Current)
        CheckBox2.Enabled = Not IsNothing(BPBS.Current)
        ComboBox1.Enabled = Not IsNothing(BPBS.Current)
        ComboBox2.Enabled = Not IsNothing(BPBS.Current)
        ComboBox3.Enabled = Not IsNothing(BPBS.Current)
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim drv As DataRowView = BPBS.AddNew()
        drv.Row.Item("cofacecoverage") = False
        drv.Row.Item("active") = False
        drv.Row.Item("bpbalance") = 0
        drv.Row.Item("bpbalanceinfc") = 0
        drv.Row.Item("creditlimit") = 0
        drv.Row.BeginEdit()
    End Sub


    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Try
            BPBS.EndEdit()
            If Me.validate Then
                Try
                    'get modified rows, send all rows to stored procedure. let the stored procedure create a new record.
                    Dim ds2 As DataSet
                    ds2 = DS.GetChanges

                    If Not IsNothing(ds2) Then
                        Dim mymessage As String = String.Empty
                        Dim ra As Integer
                        Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                        If Not DbAdapter1.BPTx(Me, mye) Then
                            DS.Merge(ds2)
                            MessageBox.Show(mye.message)
                            Exit Sub
                        End If
                        DS.Merge(ds2)
                        DS.AcceptChanges()
                        DataGridView1.Invalidate()
                        MessageBox.Show("Saved.")
                    End If
                Catch ex As Exception
                    MessageBox.Show(" Error:: " & ex.Message)
                End Try
            End If
            DataGridView1.Invalidate()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        MyBase.Validate()

        For Each drv As DataRowView In BPBS.List
            If drv.Row.RowState = DataRowState.Modified Or drv.Row.RowState = DataRowState.Added Then
                If Not validaterow(drv) Then
                    myret = False
                End If
            End If
        Next
        Return myret
    End Function

    Private Function validaterow(ByVal drv As DataRowView) As Boolean
        Dim myret As Boolean = True
        Dim sb As New StringBuilder
        If IsDBNull(drv.Row.Item("bpcode")) Then
            myret = False
            sb.Append("BP Code cannot be blank.")
        End If
        If IsDBNull(drv.Row.Item("bpname")) Then
            myret = False
            sb.Append("BP Name cannot be blank.")
        End If
        drv.Row.RowError = sb.ToString
        Return myret
    End Function

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not IsNothing(BPBS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    BPBS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub


    'Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
    '    Dim myobj As Button = CType(sender, Button)
    '    Try
    '        Select Case myobj.Name
    '            Case "Button8"
    '                Dim myform = New FormHelper(bsVendorNameHelper)
    '                myform.DataGridView1.Columns(0).DataPropertyName = "description"
    '                If myform.ShowDialog = DialogResult.OK Then
    '                    Dim drv As DataRowView = bsVendorNameHelper.Current
    '                    Dim mydrv As DataRowView = BPBS.Current
    '                    mydrv.BeginEdit()
    '                    mydrv.Row.Item("vendorcode") = drv.Row.Item("vendorcode")
    '                    mydrv.Row.Item("vendorname") = drv.Row.Item("vendorname")
    '                    mydrv.EndEdit()
    '                    'Need bellow code to sync with combobox
    '                    Dim myposition = bsVendorName.Find("vendorcode", drv.Row.Item("vendorcode"))
    '                    bsVendorName.Position = myposition
    '                End If
    '        End Select
    '    Catch ex As Exception

    '        MessageBox.Show(ex.Message)
    '    End Try

    '    DataGridView1.Invalidate()
    'End Sub


    'Private Sub ComboBox2_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted
    '    Dim myobj As ComboBox = DirectCast(sender, ComboBox)
    '    '1. Force Combobox to commit the value 
    '    For Each binding As Binding In myobj.DataBindings
    '        binding.WriteValue()
    '        binding.ReadValue()
    '    Next

    '    If Not IsNothing(BPBS.Current) Then
    '        Dim myselected1 As DataRowView = ComboBox1.SelectedItem

    '        Dim drv As DataRowView = BPBS.Current
    '        Try

    '            drv.Row.BeginEdit()
    '            drv.Row.Item("vendorname") = myselected1.Row.Item("vendorname")
    '            BPBS.EndEdit()
    '        Catch ex As Exception
    '            ComboBox1.SelectedValue = drv.Row.Item("vendorcode", DataRowVersion.Original)
    '            drv.Row.CancelEdit()
    '            MessageBox.Show(ex.Message)
    '        End Try

    '    End If
    '    DataGridView1.Invalidate()
    'End Sub

    Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged, ToolStripComboBox1.SelectedIndexChanged
        BPBS.Filter = ""
        ToolStripStatusLabel1.Text = ""
        If ToolStripTextBox1.Text <> "" And ToolStripComboBox1.SelectedIndex <> -1 Then
            Select Case ToolStripComboBox1.SelectedIndex
                Case 0
                    If Not IsNumeric(ToolStripTextBox1.Text) Then
                        ToolStripTextBox1.Select()
                        SendKeys.Send("{BACKSPACE}")
                        Exit Sub
                    End If
            End Select
            BPBS.Filter = myFields(ToolStripComboBox1.SelectedIndex).ToString & " like '%" & sender.ToString.Replace("'", "''") & "%'"
            ToolStripStatusLabel1.Text = "Record Count " & BPBS.Count
        End If
    End Sub


    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        MessageBox.Show(e.Exception.Message.ToString)
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        'Dim myform As New FormImportGroupVendor
        'myform.groupid = groupid
        'myform.ShowDialog()
        'Me.loaddata()
    End Sub


    Private Sub TextBox1_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_Validating(sender As Object, e As CancelEventArgs) Handles TextBox1.Validating
        Dim myobj = CType(sender, TextBox)
        ErrorProvider1.SetError(myobj, "")
        If myobj.Text <> "" Then
            If Not IsNumeric(myobj.Text) Then
                ErrorProvider1.SetError(myobj, "Value should be numeric.")
                e.Cancel = True
            End If
        End If        
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub TextBox3_Validating(sender As Object, e As CancelEventArgs) Handles TextBox3.Validating, TextBox5.Validating, TextBox18.Validating
        Dim myobj = CType(sender, TextBox)
        ErrorProvider1.SetError(myobj, "")
        If myobj.Text = "" Then
            myobj.Text = 0
        End If
    End Sub

    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click
        BPBS.CancelEdit()
    End Sub
End Class


Public Class CreditTerm
    Public Property CreditTerm As Object
    Public Property CreditTermName As String
    Public Sub New(_creditterm As Object, _credittermname As String)
        Me.CreditTerm = _creditterm
        Me.CreditTermName = _credittermname
    End Sub
End Class

Public Class PriceListCode
    Public Property PriceListCode As Object
    Public Property PriceListName As String
    Public Sub New(_PriceListCode As Object, _PriceListName As String)
        Me.PriceListCode = _PriceListCode
        Me.PriceListName = _PriceListName
    End Sub
End Class