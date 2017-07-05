Imports System.Data.OleDb

Public Class Form7

    'Database connection

    Private CN As OleDbConnection

    Private Sub Form7_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        'Close Connection

        CN.Close()

    End Sub

    Private Sub Form7_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        'Create Connection

        CN = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;DataSource=Stock.mdb;")

        'Open Connection

        CN.Open()

        'Load list of IDs from DB into ComboBox

        GetIDs()

        'Select first item in list

        cboStockID.SelectedIndex = 0

        'Get record from DB

        ReadRecord()

    End Sub

    'Load list of IDs from DB into ComboBox

    Private Function GetIDs() As Integer

        'Select IDs query string

        Dim strSQL As String = "SELECT StockID FROM tblItem ORDER BY StockID"

        Dim CM As New OleDbCommand(strSQL, CN)

        Dim DR As OleDbDataReader = CM.ExecuteReader

        Dim intCountRecs As Integer = 0

        'Fill ComboBox with all IDs

        Do While DR.Read

            cboStockID.Items.Add(DR.Item("StockID"))

            intCountRecs += 1

        Loop

        Return intCountRecs

    End Function

    'Input record from DB

    Private Sub ReadRecord()

        'Get selected ID from ComboBox

        Dim strStockID As String = cboStockID.Text

        'Skip if none selected

        If cboStockID.Text = "" Then Exit Sub

        'Get record with this ID

        Dim strSQL As String = "SELECT * FROM tblItem WHERE StockID = '" & strStockID & "'"

        'Get the record required from the database

        Dim CM As New OleDbCommand(strSQL, CN)

        Dim DR As OleDbDataReader = CM.ExecuteReader

        DR.Read()

        'Display fields in TextBoxes

        txtDescription.Text = DR.Item("Description").ToString

        txtPrice.Text = DR.Item("Price").ToString

        txtQuantityInStock.Text = DR.Item("QuantityInStock").ToString

        txtReorderLevel.Text = DR.Item("ReorderLevel").ToString

        txtReorderQuantity.Text = DR.Item("ReorderQuantity").ToString

        txtDateLastOrder.Text = DR.Item("DateLastOrder").ToString

        txtWhetherReceived.Text = DR.Item("WhetherReceived").ToString

    End Sub

    Private Sub cboStockID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboStockID.SelectedIndexChanged

        'If the user clears the box

        If cboStockID.Text = "" Then

            'clear all the other boxes

            ClearBoxes()

        Else

            'get the newly selected record

            ReadRecord()

        End If

    End Sub

    'Remove all displayed fields

    Private Sub ClearBoxes()

        txtDescription.Clear()

        txtPrice.Clear()

        txtQuantityInStock.Clear()

        txtReorderLevel.Clear()

        txtReorderQuantity.Clear()

        txtDateLastOrder.Clear()

        txtWhetherReceived.Clear()

    End Sub

    Private Sub UpDateIDs()

        'Memorize which row selected

        Dim intSelRow As Integer = cboStockID.SelectedIndex

        'Memorize which ID was showing

        Dim strStockID As String = cboStockID.Text

        'Empty ComboBox

        cboStockID.Items.Clear()

        'Refresh ComboBox with StockIDs

        Dim intCountRecs As Integer = GetIDs()

        'If there are any records

        If intCountRecs > 0 Then

            'If it is not a new record

            If intSelRow >= 0 Then

                'If the previously selected row doesn't exist

                If intSelRow >= cboStockID.Items.Count Then

                    'Move the selected row up

                    intSelRow -= 1

                End If

                'restore previously selected row

                cboStockID.SelectedIndex = intSelRow

            Else

                'or new row

                cboStockID.Text = strStockID

            End If

        Else

            'No records left

            ClearBoxes()

        End If

    End Sub

    Private Sub btnInsert_Click(sender As System.Object, e As System.EventArgs) Handles btnInsert.Click

        'Get new data from boxes:

        Dim strStockID As String = cboStockID.Text

        If strStockID = "" Then Exit Sub

        Dim strDescription As String = txtDescription.Text

        Dim sngPrice As Single = CSng(txtPrice.Text)

        Dim intQuantityInStock As Integer = CInt("0" & txtQuantityInStock.Text)

        Dim intReorderLevel As Integer = CInt("0" & txtReorderLevel.Text)

        Dim intReorderQuantity As Integer = CInt("0" & txtReorderQuantity.Text)

        Dim strDateLastOrder As String = CDate(txtDateLastOrder.Text)

        Dim strWhetherReceived As String = CBool(txtWhetherReceived.Text)

        'Add new record to table

        Dim strSQL As String = "INSERT INTO tblItem VALUES('" & strStockID & "','" & strDescription & "'," & sngPrice & "," & intQuantityInStock & "," & intReorderLevel & "," & intReorderQuantity & ",#" & strDateLastOrder & "#," & strWhetherReceived & ")"

        Dim CM As New OleDbCommand(strSQL, CN)

        CM.ExecuteNonQuery()

        'Refresh ComboBox

        UpDateIDs()

    End Sub

    Private Sub btnUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdate.Click

        'You cannot Update an ID

        Dim strStockID As String = cboStockID.Text

        If strStockID = "" Then Exit Sub

        'Get new data from boxes:

        Dim strDescription As String = txtDescription.Text

        Dim sngPrice As Single = CSng(txtPrice.Text)

        Dim intQuantityInStock As Integer = CInt("0" & txtQuantityInStock.Text)

        Dim intReorderLevel As Integer = CInt("0" & txtReorderLevel.Text)

        Dim intReorderQuantity As Integer = CInt("0" & txtReorderQuantity.Text)

        Dim strDateLastOrder As String = CDate(txtDateLastOrder.Text)

        Dim strWhetherReceived As String = CBool(txtWhetherReceived.Text)

        Dim strSQL As String = "UPDATE tblItem SET " & "Description='" & strDescription & "'," & "Price=" & sngPrice & "," & "QuantityInStock=" & intQuantityInStock & "," & "ReorderLevel=" & intReorderLevel & "," & "ReorderQuantity=" & intReorderQuantity & "," & "DateLastOrder=#" & strDateLastOrder & "#," & "WhetherReceived=" & strWhetherReceived & " " & "WHERE StockID = '" & strStockID & "'"

        Dim CM As New OleDbCommand(strSQL, CN)

        CM.ExecuteNonQuery()

    End Sub

    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click

        Dim strStockID As String = cboStockID.Text

        If strStockID = "" Then Exit Sub

        Dim strSQL As String = "DELETE FROM tblItem " & "WHERE StockID = '" & strStockID & "'"

        Dim CM As New OleDbCommand(strSQL, CN)

        CM.ExecuteNonQuery()

        'Refresh ComboBox

        UpDateIDs()

    End Sub

End Class