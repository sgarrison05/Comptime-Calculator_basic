'Title                  Comptime Calculator
'Purpose                To calculate comptime time earned or spent
'                       in a particular instance
'Created By             Shon Garrison, December 2008
'Updated Last           September 2009

Public Class frm_About

    Private Sub okButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles okButton.Click
        Me.Close()
    End Sub

    Private Sub aboutForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.copywriteTextBox.ReadOnly = True
        Me.deptidTextBox.ReadOnly = True

    End Sub

   
End Class