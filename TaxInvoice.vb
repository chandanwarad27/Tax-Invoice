Imports C1.Win.C1FlexGrid
Imports System.Data.SqlClient


Public Class TaxInvoice
    Private Sub vsfTaxInvoice_BeforeEdit(sender As Object, e As RowColEventArgs) Handles vsfTaxInvoice.BeforeEdit
        Vsf_RowCountChanged(e.Row)
        If e.Col = 2 Then
            vsfTaxInvoice.Cols(e.Col).DataMap = LoadItemCode()
        End If
    End Sub

    Private Function LoadItemCode()
        Dim query As String
        Dim reader As SqlDataReader

        Try
            query = "SELECT * From ItemDetails"
            Dim conn As New ConnectToDB
            reader = conn.GetConnection(query)
            Dim itemCode As New Specialized.ListDictionary
            While reader.Read
                itemCode.Add(reader.Item(0), reader.Item(2))
            End While
            conn.sqlConn.Close()
            Return itemCode
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Private Sub Vsf_RowCountChanged(index As Integer)
        For i = 0 To vsfTaxInvoice.Rows.Count - 1
            vsfTaxInvoice.SetData(index, 1, index)
        Next
    End Sub

    Private Sub vsfTaxInvoice_KeyPress(sender As Object, e As KeyPressEventArgs) Handles vsfTaxInvoice.KeyPress
        If e.KeyChar = vbTab Then
            vsfTaxInvoice.Rows.Add()
        End If
    End Sub

    Private Sub TaxInvoice_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        vsfTaxInvoice.Rows.Add()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim query As String
        Dim conn As New ConnectToDB()
        Dim reader As SqlDataReader

        For index As Integer = 1 To vsfTaxInvoice.Rows.Count - 1 Step 1
            Try
                Dim dialog As DialogResult
                dialog = MessageBox.Show("Are you sure you want to Save", "Save", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                If dialog = DialogResult.Yes Then
                    query = "Insert Into Invoice(ItemDetailsId,ItemDescription,Quantity,Rate,Amount) Values ('" & vsfTaxInvoice.Rows(index).Item(2) & "','" & vsfTaxInvoice.Rows(index).Item(5) & "','" & vsfTaxInvoice.Rows(index).Item(6) & "','" & vsfTaxInvoice.Rows(index).Item(7) & "','" & vsfTaxInvoice.Rows(index).Item(8) & "')"
                    reader = conn.GetConnection(query)
                    conn.reader.Close()

                        MessageBox.Show("Data saved successfully", "Successfull !")

                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Next
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Dim dialog As DialogResult
        dialog = MessageBox.Show("Are you sure you want to delete", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If dialog = DialogResult.Yes Then
            Dim selectedItem = Selected_Item()

            For Each item In selectedItem
                vsfTaxInvoice.RemoveItem(item)
            Next
        End If
    End Sub

    Private Function Selected_Item()
        Dim items As New List(Of Integer)
        For index = 1 To vsfTaxInvoice.Rows.Count - 1
            If (vsfTaxInvoice.GetData(index, 0)) = "True" Then
                items.Add(vsfTaxInvoice.GetData(index, 1))
            End If
        Next
        Return items
    End Function

    Private Sub vsfTaxInvoice_AfterEdit(sender As Object, e As RowColEventArgs) Handles vsfTaxInvoice.AfterEdit
        Vsf_RowCountChanged(e.Row)

        If e.Col = 2 Then
            Try
                Dim query As String
                Dim conn As New ConnectToDB()
                Dim reader As SqlDataReader

                query = "Select * From ItemDetails where ItemDetailsId = '" & vsfTaxInvoice.GetData(e.Row, 2) & "'"

                reader = conn.GetConnection(query)

                While reader.Read
                    vsfTaxInvoice.SetData(e.Row, 3, reader.Item(2))
                    vsfTaxInvoice.SetData(e.Row, 4, reader.Item(3))
                End While
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If

        If e.Col = 6 Or e.Col = 7 Then
            If vsfTaxInvoice.GetData(e.Row, 6) <> Nothing And vsfTaxInvoice.GetData(e.Row, 6) <> Nothing Then
                vsfTaxInvoice.SetData(e.Row, 8, (vsfTaxInvoice.GetData(e.Row, 6) * vsfTaxInvoice.GetData(e.Row, 7)))
            End If
        End If
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Dim dialog As DialogResult
        dialog = MessageBox.Show("Are you sure you want to Exit", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If dialog = DialogResult.Yes Then
            End
        End If
    End Sub

    Private Sub vsfTaxInvoice_ValidateEdit(sender As Object, e As ValidateEditEventArgs) Handles vsfTaxInvoice.ValidateEdit
        With vsfTaxInvoice
            Select Case e.Col
                Case 6
                    If Not IsNumeric(.Editor.Text) And .Editor.Text <> Nothing Then
                        MessageBox.Show("Enter numeric value", "Error !", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        e.Cancel = True
                    End If
                Case 7
                    If Not IsNumeric(.Editor.Text) And .Editor.Text <> Nothing Then
                        MessageBox.Show("Enter numeric value", "Error !", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        e.Cancel = True
                    End If
            End Select
        End With
    End Sub
End Class
