Public Class frmMain

    Friend persister As PFW.CSIST203.Project2.Persisters.Excel.ExcelPersister

    Private Sub FrmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' TODO: Implement
    End Sub

    Private Sub BtnPrevious_Click(sender As Object, e As EventArgs) Handles btnPrevious.Click
        Throw New NotImplementedException()
    End Sub

    Private Sub BtnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        Throw New NotImplementedException()
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        MyBase.OnFormClosing(e)
        ' TODO: Implement
    End Sub

    ''' <summary>
    ''' Handle the File -> Open dialog box used for selecting the excel file that is utilized by the front-end
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        Dim result = OpenFileDialog.ShowDialog()
        If result = DialogResult.OK Then
            persister.Dispose()
            persister = New PFW.CSIST203.Project2.Persisters.Excel.ExcelPersister(OpenFileDialog.FileName)
            txtRow.Text = "0" ' reset back to zero
            txtFirstname.Text = String.Empty
            txtLastname.Text = String.Empty
            txtEmailAddress.Text = String.Empty
            txtBusinessPhone.Text = String.Empty
            txtCompany.Text = String.Empty
            txtTitle.Text = String.Empty
        End If
    End Sub

End Class
