Public Class frm_Reconcillation

    Private Sub btn_Preview_Click(sender As System.Object, e As System.EventArgs) Handles btn_Preview.Click

        Dim objectreader As New System.IO.StreamReader("C:\Comptime\comptimerun.txt")


        If My.Computer.FileSystem.FileExists("C:\Comptime\comptimerun.txt") Then


            Me.txt_ReconPrev.Text = objectreader.ReadToEnd

            objectreader.Close()

        Else : MsgBox("Error...File Does not exist!")


        End If

    End Sub

    Private Sub btn_Return_Click(sender As System.Object, e As System.EventArgs) Handles btn_Return.Click

        Me.Close()
        frm_Main.Show()

    End Sub

    Private Sub btn_Clear_Click(sender As System.Object, e As System.EventArgs) Handles btn_Clear.Click


        Me.txt_ReconPrev.Text = String.Empty

    End Sub
End Class