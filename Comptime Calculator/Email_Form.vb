Imports System.Net.Mail


Public Class frm_Email

    Dim SMTPClient As New SmtpClient
    Dim emailChoice As String = ""
    Dim ESubject As String = "Emailing Comptime Data for " & DatePart(DateInterval.Month, Today) & "/" & _
                                                                        DatePart(DateInterval.Year, Today)

    Private Sub btn_Email_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Email.Click

        Dim username As String = txt_Username.Text
        Dim pw As String = txt_PW.Text
        Dim pathAttach As String = txt_Path.Text
        Dim today As Date = Date.Now

        emailChoice = Me.cboxEmailProvider.Text

        Select Case emailChoice

            Case "County/Local"

                SMTPClient.Host = "orangeco1.co.orange.tx.us"
                SMTPClient.Port = 24

            Case "Google"
                SMTPClient.Host = "smtp.gmail.com"
                SMTPClient.Port = 587
                SMTPClient.EnableSsl = True

            Case "AT&T"
                SMTPClient.Host = "outbound.att.net"
                SMTPClient.Port = 587
                SMTPClient.EnableSsl = True

            Case "Yahoo"
                SMTPClient.Host = "smtp.mail.yahoo.com"
                SMTPClient.Port = 587
                SMTPClient.EnableSsl = True

            Case "MSN"
                SMTPClient.Host = "smtp.live.com"
                SMTPClient.Port = 587
                SMTPClient.EnableSsl = True

        End Select

        If Me.txt_Path.Text = "" Or Me.txt_Username.Text = "" Or Me.txt_PW.Text = "" Or Me.txt_From.Text = "" Or Me.Cmb_txt.Text = "" Or _
            Me.txt_Subject.Text = "" Or Me.txt_Msg.Text = "" Then

            MessageBox.Show("You forgot something or to attach your comtimerun.txt file. Please check all blanks.", "Important", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Me.brn_Browse.Focus()

        Else

            Dim UsernamePassword As New Net.NetworkCredential(username, _
                                                                        pw)
            SMTPClient.Credentials = UsernamePassword

            Dim fromSender As String = txt_From.Text
            Dim toReceipiant As String = Cmb_txt.Text
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
        End If


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

        Cmb_txt.Text = ""
        Me.txt_Path.Text = ""
        Me.txt_Msg.Text = ""
        Me.txt_Username.Focus()


    End Sub

    Private Sub frm_Email_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        'Fills the sender box
        Cmb_txt.Items.Add("[Enter/Select One]")
        'Cmb_txt.Items.Add("ltaylor@co.orange.tx.us")
        Cmb_txt.Items.Add("wcolbert@co.orange.tx.us")
        Cmb_txt.Items.Add("ccorder@co.orange.tx.us")

        'Fills the Email Provider Box
        Me.cboxEmailProvider.Items.Add("[Choose Provider]")
        Me.cboxEmailProvider.Items.Add("County/Local")
        Me.cboxEmailProvider.Items.Add("Google")
        Me.cboxEmailProvider.Items.Add("Yahoo")
        Me.cboxEmailProvider.Items.Add("AT&T")
        Me.cboxEmailProvider.Items.Add("MSN")

        'sets default email selection
        Cmb_txt.SelectedIndex = 0
        cboxEmailProvider.SelectedIndex = 0


        txt_Subject.Text = ESubject

    End Sub


    Private Sub cboxEmailProvider_TextChanged(sender As Object, e As EventArgs) Handles cboxEmailProvider.TextChanged

        emailChoice = Me.cboxEmailProvider.Text

        Select Case emailChoice

            Case "County/Local"

                SmtpClient.Host = "orangeco1.co.orange.tx.us"
                SmtpClient.Port = 24

            Case "Google"
                SmtpClient.Host = "smtp.gmail.com"
                SMTPClient.Port = 587
                SmtpClient.EnableSsl = True

            Case "AT&T"
                SmtpClient.Host = "outbound.att.net"
                SMTPClient.Port = 587
                SmtpClient.EnableSsl = True

            Case "Yahoo"
                SmtpClient.Host = "smtp.mail.yahoo.com"
                SMTPClient.Port = 587
                SmtpClient.EnableSsl = True

            Case "MSN"
                SmtpClient.Host = "smtp.live.com"
                SMTPClient.Port = 587
                SmtpClient.EnableSsl = True

        End Select

        'clears labels
        Me.txt_Username.Text = ""
        Me.txt_From.Text = ""
        Me.txt_PW.Text = ""



    End Sub
End Class