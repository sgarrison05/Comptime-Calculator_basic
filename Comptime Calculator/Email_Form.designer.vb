<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_Email
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_Email))
        Me.btn_Exit = New System.Windows.Forms.Button()
        Me.btn_Clear = New System.Windows.Forms.Button()
        Me.btn_Email = New System.Windows.Forms.Button()
        Me.txt_From = New System.Windows.Forms.TextBox()
        Me.txt_Path = New System.Windows.Forms.TextBox()
        Me.brn_Browse = New System.Windows.Forms.Button()
        Me.lbl_From = New System.Windows.Forms.Label()
        Me.lbl_To = New System.Windows.Forms.Label()
        Me.lbl_Path = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.txt_Subject = New System.Windows.Forms.TextBox()
        Me.txt_Msg = New System.Windows.Forms.TextBox()
        Me.lbl_Subject = New System.Windows.Forms.Label()
        Me.lbl_Msg = New System.Windows.Forms.Label()
        Me.txt_PW = New System.Windows.Forms.TextBox()
        Me.lbl_PW = New System.Windows.Forms.Label()
        Me.Cmb_txt = New System.Windows.Forms.ComboBox()
        Me.cboxEmailProvider = New System.Windows.Forms.ComboBox()
        Me.lblEmailProvider = New System.Windows.Forms.Label()
        Me.txt_Username = New System.Windows.Forms.TextBox()
        Me.lbl_Username = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btn_Exit
        '
        Me.btn_Exit.DialogResult = System.Windows.Forms.DialogResult.Yes
        Me.btn_Exit.Location = New System.Drawing.Point(14, 583)
        Me.btn_Exit.Name = "btn_Exit"
        Me.btn_Exit.Size = New System.Drawing.Size(83, 27)
        Me.btn_Exit.TabIndex = 8
        Me.btn_Exit.Text = "Rtn to Main"
        Me.btn_Exit.UseVisualStyleBackColor = True
        '
        'btn_Clear
        '
        Me.btn_Clear.Location = New System.Drawing.Point(182, 583)
        Me.btn_Clear.Name = "btn_Clear"
        Me.btn_Clear.Size = New System.Drawing.Size(83, 27)
        Me.btn_Clear.TabIndex = 9
        Me.btn_Clear.Text = "Clear Form"
        Me.btn_Clear.UseVisualStyleBackColor = True
        '
        'btn_Email
        '
        Me.btn_Email.Location = New System.Drawing.Point(365, 583)
        Me.btn_Email.Name = "btn_Email"
        Me.btn_Email.Size = New System.Drawing.Size(83, 27)
        Me.btn_Email.TabIndex = 10
        Me.btn_Email.Text = "Email Data"
        Me.btn_Email.UseVisualStyleBackColor = True
        '
        'txt_From
        '
        Me.txt_From.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_From.Location = New System.Drawing.Point(10, 209)
        Me.txt_From.Multiline = True
        Me.txt_From.Name = "txt_From"
        Me.txt_From.Size = New System.Drawing.Size(206, 29)
        Me.txt_From.TabIndex = 2
        '
        'txt_Path
        '
        Me.txt_Path.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_Path.Location = New System.Drawing.Point(10, 496)
        Me.txt_Path.Multiline = True
        Me.txt_Path.Name = "txt_Path"
        Me.txt_Path.Size = New System.Drawing.Size(443, 29)
        Me.txt_Path.TabIndex = 6
        '
        'brn_Browse
        '
        Me.brn_Browse.Location = New System.Drawing.Point(15, 531)
        Me.brn_Browse.Name = "brn_Browse"
        Me.brn_Browse.Size = New System.Drawing.Size(83, 27)
        Me.brn_Browse.TabIndex = 7
        Me.brn_Browse.Text = "Browse"
        Me.brn_Browse.UseVisualStyleBackColor = True
        '
        'lbl_From
        '
        Me.lbl_From.AutoSize = True
        Me.lbl_From.Location = New System.Drawing.Point(7, 193)
        Me.lbl_From.Name = "lbl_From"
        Me.lbl_From.Size = New System.Drawing.Size(124, 13)
        Me.lbl_From.TabIndex = 7
        Me.lbl_From.Text = "Email Address (Sending):"
        '
        'lbl_To
        '
        Me.lbl_To.AutoSize = True
        Me.lbl_To.Location = New System.Drawing.Point(244, 193)
        Me.lbl_To.Name = "lbl_To"
        Me.lbl_To.Size = New System.Drawing.Size(130, 13)
        Me.lbl_To.TabIndex = 8
        Me.lbl_To.Text = "Email Address(Receiving):"
        '
        'lbl_Path
        '
        Me.lbl_Path.AutoSize = True
        Me.lbl_Path.CausesValidation = False
        Me.lbl_Path.Location = New System.Drawing.Point(8, 480)
        Me.lbl_Path.Name = "lbl_Path"
        Me.lbl_Path.Size = New System.Drawing.Size(182, 13)
        Me.lbl_Path.TabIndex = 9
        Me.lbl_Path.Text = "Path to comptimerun.txt file to attach:"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "Select Your File"
        Me.OpenFileDialog1.Filter = "Text File(*.txt)|*.txt|All Files|*.*"
        Me.OpenFileDialog1.InitialDirectory = "C:\Comptime\"
        '
        'txt_Subject
        '
        Me.txt_Subject.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_Subject.Location = New System.Drawing.Point(10, 267)
        Me.txt_Subject.Multiline = True
        Me.txt_Subject.Name = "txt_Subject"
        Me.txt_Subject.Size = New System.Drawing.Size(443, 29)
        Me.txt_Subject.TabIndex = 4
        '
        'txt_Msg
        '
        Me.txt_Msg.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_Msg.Location = New System.Drawing.Point(10, 317)
        Me.txt_Msg.Multiline = True
        Me.txt_Msg.Name = "txt_Msg"
        Me.txt_Msg.Size = New System.Drawing.Size(443, 149)
        Me.txt_Msg.TabIndex = 5
        '
        'lbl_Subject
        '
        Me.lbl_Subject.AutoSize = True
        Me.lbl_Subject.Location = New System.Drawing.Point(7, 251)
        Me.lbl_Subject.Name = "lbl_Subject"
        Me.lbl_Subject.Size = New System.Drawing.Size(46, 13)
        Me.lbl_Subject.TabIndex = 12
        Me.lbl_Subject.Text = "Subject:"
        '
        'lbl_Msg
        '
        Me.lbl_Msg.AutoSize = True
        Me.lbl_Msg.Location = New System.Drawing.Point(7, 301)
        Me.lbl_Msg.Name = "lbl_Msg"
        Me.lbl_Msg.Size = New System.Drawing.Size(77, 13)
        Me.lbl_Msg.TabIndex = 13
        Me.lbl_Msg.Text = "Brief Message:"
        '
        'txt_PW
        '
        Me.txt_PW.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_PW.Location = New System.Drawing.Point(10, 152)
        Me.txt_PW.Multiline = True
        Me.txt_PW.Name = "txt_PW"
        Me.txt_PW.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txt_PW.Size = New System.Drawing.Size(206, 29)
        Me.txt_PW.TabIndex = 1
        '
        'lbl_PW
        '
        Me.lbl_PW.AutoSize = True
        Me.lbl_PW.Location = New System.Drawing.Point(8, 136)
        Me.lbl_PW.Name = "lbl_PW"
        Me.lbl_PW.Size = New System.Drawing.Size(56, 13)
        Me.lbl_PW.TabIndex = 16
        Me.lbl_PW.Text = "Password:"
        '
        'Cmb_txt
        '
        Me.Cmb_txt.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cmb_txt.FormattingEnabled = True
        Me.Cmb_txt.Location = New System.Drawing.Point(247, 209)
        Me.Cmb_txt.Name = "Cmb_txt"
        Me.Cmb_txt.Size = New System.Drawing.Size(207, 28)
        Me.Cmb_txt.TabIndex = 3
        '
        'cboxEmailProvider
        '
        Me.cboxEmailProvider.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboxEmailProvider.FormattingEnabled = True
        Me.cboxEmailProvider.Location = New System.Drawing.Point(10, 29)
        Me.cboxEmailProvider.Name = "cboxEmailProvider"
        Me.cboxEmailProvider.Size = New System.Drawing.Size(155, 28)
        Me.cboxEmailProvider.TabIndex = 18
        '
        'lblEmailProvider
        '
        Me.lblEmailProvider.AutoSize = True
        Me.lblEmailProvider.Location = New System.Drawing.Point(8, 13)
        Me.lblEmailProvider.Name = "lblEmailProvider"
        Me.lblEmailProvider.Size = New System.Drawing.Size(77, 13)
        Me.lblEmailProvider.TabIndex = 19
        Me.lblEmailProvider.Text = "Email Provider:"
        '
        'txt_Username
        '
        Me.txt_Username.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_Username.Location = New System.Drawing.Point(9, 92)
        Me.txt_Username.Multiline = True
        Me.txt_Username.Name = "txt_Username"
        Me.txt_Username.Size = New System.Drawing.Size(206, 29)
        Me.txt_Username.TabIndex = 0
        '
        'lbl_Username
        '
        Me.lbl_Username.AutoSize = True
        Me.lbl_Username.Location = New System.Drawing.Point(7, 76)
        Me.lbl_Username.Name = "lbl_Username"
        Me.lbl_Username.Size = New System.Drawing.Size(228, 13)
        Me.lbl_Username.TabIndex = 17
        Me.lbl_Username.Text = "Username, Email Address, or Email Credentials:"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(268, 12)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(164, 168)
        Me.PictureBox1.TabIndex = 20
        Me.PictureBox1.TabStop = False
        '
        'frm_Email
        '
        Me.AcceptButton = Me.btn_Email
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btn_Exit
        Me.ClientSize = New System.Drawing.Size(464, 622)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.lblEmailProvider)
        Me.Controls.Add(Me.cboxEmailProvider)
        Me.Controls.Add(Me.Cmb_txt)
        Me.Controls.Add(Me.lbl_Username)
        Me.Controls.Add(Me.lbl_PW)
        Me.Controls.Add(Me.txt_PW)
        Me.Controls.Add(Me.txt_Username)
        Me.Controls.Add(Me.lbl_Msg)
        Me.Controls.Add(Me.lbl_Subject)
        Me.Controls.Add(Me.txt_Msg)
        Me.Controls.Add(Me.txt_Subject)
        Me.Controls.Add(Me.lbl_Path)
        Me.Controls.Add(Me.lbl_To)
        Me.Controls.Add(Me.lbl_From)
        Me.Controls.Add(Me.brn_Browse)
        Me.Controls.Add(Me.txt_Path)
        Me.Controls.Add(Me.txt_From)
        Me.Controls.Add(Me.btn_Email)
        Me.Controls.Add(Me.btn_Clear)
        Me.Controls.Add(Me.btn_Exit)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frm_Email"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Send Data by Email"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btn_Exit As System.Windows.Forms.Button
    Friend WithEvents btn_Clear As System.Windows.Forms.Button
    Friend WithEvents btn_Email As System.Windows.Forms.Button
    Friend WithEvents txt_From As System.Windows.Forms.TextBox
    Friend WithEvents txt_Path As System.Windows.Forms.TextBox
    Friend WithEvents brn_Browse As System.Windows.Forms.Button
    Friend WithEvents lbl_To As System.Windows.Forms.Label
    Friend WithEvents lbl_Path As System.Windows.Forms.Label
    Friend WithEvents lbl_From As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents txt_Subject As System.Windows.Forms.TextBox
    Friend WithEvents txt_Msg As System.Windows.Forms.TextBox
    Friend WithEvents lbl_Subject As System.Windows.Forms.Label
    Friend WithEvents lbl_Msg As System.Windows.Forms.Label
    Friend WithEvents txt_PW As System.Windows.Forms.TextBox
    Friend WithEvents lbl_PW As System.Windows.Forms.Label
    Friend WithEvents Cmb_txt As System.Windows.Forms.ComboBox
    Friend WithEvents cboxEmailProvider As System.Windows.Forms.ComboBox
    Friend WithEvents lblEmailProvider As System.Windows.Forms.Label
    Friend WithEvents txt_Username As System.Windows.Forms.TextBox
    Friend WithEvents lbl_Username As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
End Class
