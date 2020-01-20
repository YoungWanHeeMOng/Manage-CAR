
Imports System.Windows.Forms
Imports System.IO
Imports Microsoft.Win32
Public Class subPartsList

    Private Sub subGrid_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles subGrid.CellContentClick

    End Sub

    Private Sub subGrid_CellMouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles subGrid.CellMouseDoubleClick
        '주유소, 주소를 메인으로 보내준다
        Dim i As Integer
        i = Me.subGrid.CurrentRow.Index

        Main.gbtTx_Store.Text = Me.subGrid.Rows(i).Cells(0).Value
        Main.gbtTx_Etc.Text = Me.subGrid.Rows(i).Cells(1).Value

        Hide()
    End Sub

    Private Sub subGrid1_CellMouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles subGrid1.CellMouseDoubleClick
        '주유소, 주소를 메인으로 보내준다
        Dim i As Integer
        i = Me.subGrid1.CurrentRow.Index

        Main.gbtTx_Store.Text = Me.subGrid1.Rows(i).Cells(0).Value
        Main.gbtTx_Etc.Text = Me.subGrid1.Rows(i).Cells(1).Value

        Hide()
    End Sub

    Private Sub txStore_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txStore.KeyDown
        Me.txEtc.Text = ""
    End Sub

    Private Sub txStore_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txStore.TextChanged
        srchPartsStore()
    End Sub

    Private Sub btnInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInput.Click
        '정비소와 주소 옮기기

        Main.gbtTx_Store.Text = Me.txStore.Text
        Main.gbtTx_Etc.Text = Me.txEtc.Text

        movePartsAddress()

        Me.Hide()
    End Sub

    Private Sub subPartsList_Load(sender As Object, e As EventArgs) Handles Me.Load

        Me.txStore.ImeMode = ImeMode.Hangul
        Me.txEtc.ImeMode = ImeMode.Hangul

    End Sub
End Class