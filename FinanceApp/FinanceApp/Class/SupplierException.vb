Imports System.Text

Public Class SupplierException
    Private MainQuery As StringBuilder
    Public Sub New()
        MainQuery = New StringBuilder
        MainQuery.Append("with vi as (select invoicenumber ,sum(iws) as amount from finance.vendorinvoice group by invoicenumber)," &
                         " sapi as (select reference2,sum(amount1) * -1 as amount from finance.sapinvoice group by reference2)," &
                         " match as (select sapi.* from sapi" &
                         " inner join vi on vi.invoicenumber = sapi.reference2 and sapi.amount = vi.amount)," &
                         " viexception as (select * from vi except select * from match)," &
                         " sapiexception as (select * from vi except select * from match)," &
                         " rptvendorexception as (select fv.invoicenumber,fv.month,fv.mydate,fv.iws,fv.gl,fv.dc,fv.cpy from finance.vendorinvoice fv inner join viexception ve on ve.invoicenumber = fv.invoicenumber order by ve.invoicenumber)," &
                         " rptsapexception as (select fs.stat,fs.vendor,fs.documentno,fs.myreference,fs.mytype,fs.sg,fs.clrngcode,fs.username,fs.remarks,fs.docdate,fs.entrydate,fs.postingdate,fs.clearingdate,fs.crcy1,fs.amount1,fs.amount2,fs.crcy2,fs.reference2 from finance.sapinvoice fs inner join viexception ve on ve.invoicenumber = fs.reference2 order by reference2,myreference,documentno)," &
                         " rptsapmatch as (select fs.stat,fs.vendor,fs.documentno,fs.myreference,fs.mytype,fs.sg,fs.clrngcode,fs.username,fs.remarks,fs.docdate,fs.entrydate,fs.postingdate,fs.clearingdate,fs.crcy1,fs.amount1,fs.amount2,fs.crcy2,fs.reference2 from finance.sapinvoice fs inner join match m on m.reference2 = fs.reference2 order by reference2,myreference,documentno)," &
                         " rptvendormatch as (select fv.invoicenumber,fv.month,fv.mydate,fv.iws,fv.gl,fv.dc,fv.cpy from finance.vendorinvoice fv inner join match m on m.reference2 = fv.invoicenumber order by reference2)," &
                         " rptsapall as (select fs.stat,fs.vendor,fs.documentno,fs.myreference,fs.mytype,fs.sg,fs.clrngcode,fs.username,fs.remarks,fs.docdate,fs.entrydate,fs.postingdate,fs.clearingdate,fs.crcy1,fs.amount1,fs.amount2,fs.crcy2,fs.reference2 from finance.sapinvoice fs order by reference2,myreference,documentno)")
    End Sub



    Public Sub RunSupplierException()
        Dim filename As String = "SupplierException-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
        MainQuery.Append(" select * from rptvendorexception;")
        ExportToExcelFile.ExportToExcelAskDirectory(filename, MainQuery.ToString)
    End Sub
    Public Sub RunSAPException()
        Dim filename As String = "SAPException-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
        MainQuery.Append(" select * from rptsapexception;")
        ExportToExcelFile.ExportToExcelAskDirectory(filename, MainQuery.ToString)
    End Sub
    Public Sub RunSupplierMatched()
        Dim filename As String = "SupplierMatched-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
        MainQuery.Append(" select * from rptvendormatch;")
        ExportToExcelFile.ExportToExcelAskDirectory(filename, MainQuery.ToString)
    End Sub
    Public Sub RunSAPMatched()
        Dim filename As String = "SAPMatched-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
        MainQuery.Append(" select * from rptsapmatch;")
        ExportToExcelFile.ExportToExcelAskDirectory(filename, MainQuery.ToString)
    End Sub
    Public Sub RunSAPRawData()
        Dim filename As String = "SAPRawData-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
        MainQuery.Append(" select * from rptsapall;")
        ExportToExcelFile.ExportToExcelAskDirectory(filename, MainQuery.ToString)
    End Sub
End Class
