Imports Microsoft.Office.Interop.Excel
Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim oConNewList As New Odbc.OdbcConnection
        Dim oConSupList As New Odbc.OdbcConnection

        Dim oDataAda As Odbc.OdbcDataAdapter

        Dim oExcel As New Microsoft.Office.Interop.Excel.Application

        oConNewList.ConnectionString = "Dsn=Work;dbq=C:\Test\work\MYproducts122908.xlsx;defaultdir=C:\Test\work;driverid=1046;fil=excel 12.0;maxbuffersize=2048;pagetimeout=5"
        oConNewList.Open()




        Dim oWB As Workbook = oExcel.Workbooks.Open("C:\Test\work\Suppliers.xls")



        Dim oWS As New Worksheet
        Dim oMyWs As New Worksheet

        Dim ProdCode As String


        Dim ds As New DataSet

        Dim oCommand As New Odbc.OdbcCommand
        oCommand.Connection = oConNewList

        oWS = oWB.Sheets(1)

        Dim i As Integer

        oDataAda = New Data.Odbc.OdbcDataAdapter("select * from [products122908]", oConNewList)
        oDataAda.Fill(ds)

        oConNewList.Close()
        oConNewList.Dispose()

        Dim dr() As DataRow

        Dim oMyWB As Workbook = oExcel.Workbooks.Open("C:\Test\work\MYproducts122908.xlsx")
        oMyWs = oMyWB.Sheets(1)


        oMyWs.Cells(oMyWs.Range("A65536").End(XlDirection.xlUp).Row, 1).select()

        oExcel.Visible = True

        Dim CurrentRow As Integer

        CurrentRow = oExcel.ActiveCell.Row + 1



        

        For i = 2 To 65536            
            ProdCode = oWS.Range("A" + i.ToString().Trim()).Value
            If ProdCode.Trim() = "" Then Exit For
            dr = ds.Tables(0).Select("SKU='" + ProdCode + "'")
            If dr.Count= 0 Then
                oMyWs.Cells(CurrentRow, 1).value = ProdCode
                oMyWs.Cells(CurrentRow, 2).value = oWS.Range("B" + i.ToString().Trim()).Value
                oMyWs.Cells(CurrentRow, 3).value = "basic"
                oMyWs.Cells(CurrentRow, 5).value = 0
                oMyWs.Cells(CurrentRow, 6).value = oWS.Range("B" + i.ToString().Trim()).Value
                oMyWs.Cells(CurrentRow, 7).value = "$7.50 ground shipping.  Need Help? Email us at sales@FastFittings.com"
                oMyWs.Cells(CurrentRow, 8).value = "Free ground shipping on orders over $99.00"                
                CurrentRow = CurrentRow + 1
            End If
            dr = Nothing

        Next

        MessageBox.Show("Finished")

        oExcel.Quit()






        oConNewList.Close()
    End Sub
End Class
