Imports System.Net.Mail


Public Class frm_Email

    Dim ESubject As String = "Emailing Comptime Data for " & DatePart(DateInterval.Month, Today) & "/" & _
                                                                        DatePart(DateInterval.Year, Today)

    Private Sub btn_Email_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Email.Click

        Try

        Catch ex As Exception

        End Try

        Dim username As String = txt_Username.Text
        Dim pw As String = txt_PW.Text
        Dim pathAttach As String = txt_Path.Text
        Dim today As Date = Date.Now
        Dim SMTPClient As New SmtpClient

        SMTPClient.Host = "orangeco1.co.orange.tx.us"
        SMTPClient.Port = 24

        Dim UsernamePassword As New Net.NetworkCredential(username, _
                                                            pw)
        SMTPClient.Credentials = UsernamePassword

        Dim fromSender As String = txt_From.Text
        Dim toReceipiant As String = txt_To.Text

        Dim MsgBody As String = txt_Msg.Text

        Dim MailMsg As New MailMessage(fromSender, _
                                        toReceipiant, _
                                        ESubject, _
                                        MsgBody)
        Dim MsgAtt As New Attachment(pathAttach)

        MailMsg.Attachments.Add(MsgAtt)

        Try
            SMTPClient.Send(MailMsg)
            MsgBox("Email Sent Successfully!")

            Call clearlabels()

            Me.Close()
            frm_Main.Show()


        Catch ex As Exception
            MsgBox("Error...Something went wrong! Please Try again later.")

        End Try
        

    End Sub

    Private Sub btn_Exit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Exit.Click
        Me.Close()
        frm_Main.Show()


    End Sub


    Private Sub brn_Browse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles brn_Browse.Click
        Me.OpenFileDialog1.ShowDialog()

    End Sub

    Private Sub OpenFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        txt_Path.Text = OpenFileDialog1.FileName

    End Sub

    Private Sub btn_Clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Clear.Click

        Call clearlabels()


    End Sub

    Public Sub clearlabels()

        Me.txt_To.Text = ""
        Me.txt_Path.Text = ""
        Me.txt_Msg.Text = ""
        Me.txt_Username.Focus()


    End Sub

    Private Sub frm_Email_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        txt_Subject.Text = ESubject

    End Sub
End Class