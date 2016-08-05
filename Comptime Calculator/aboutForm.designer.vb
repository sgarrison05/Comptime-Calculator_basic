<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_About
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_About))
        Me.copywriteTextBox = New System.Windows.Forms.TextBox()
        Me.deptidTextBox = New System.Windows.Forms.TextBox()
        Me.ccalcPictureBox = New System.Windows.Forms.PictureBox()
        Me.okButton = New System.Windows.Forms.Button()
        CType(Me.ccalcPictureBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'copywriteTextBox
        '
        Me.copywriteTextBox.BackColor = System.Drawing.SystemColors.MenuBar
        Me.copywriteTextBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.copywriteTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.copywriteTextBox.Location = New System.Drawing.Point(156, 37)
        Me.copywriteTextBox.Multiline = True
        Me.copywriteTextBox.Name = "copywriteTextBox"
        Me.copywriteTextBox.Size = New System.Drawing.Size(233, 66)
        Me.copywriteTextBox.TabIndex = 0
        Me.copywriteTextBox.Text = "Comptime Calculator" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Version: 4.10" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Created By:  Shon Garrison 2009" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Last Updated" & _
    ": August 2016"
        '
        'deptidTextBox
        '
        Me.deptidTextBox.BackColor = System.Drawing.SystemColors.MenuBar
        Me.deptidTextBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.deptidTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.deptidTextBox.Location = New System.Drawing.Point(156, 126)
        Me.deptidTextBox.Multiline = True
        Me.deptidTextBox.Name = "deptidTextBox"
        Me.deptidTextBox.Size = New System.Drawing.Size(280, 126)
        Me.deptidTextBox.TabIndex = 1
        Me.deptidTextBox.Text = "For use by:" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Orange County Juvenile Probation Dept." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "213 Market St." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Orange, Texa" & _
    "s 77630" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(409) 882 - 7885" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(409) 882 - 7844"
        '
        'ccalcPictureBox
        '
        Me.ccalcPictureBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ccalcPictureBox.Image = CType(resources.GetObject("ccalcPictureBox.Image"), System.Drawing.Image)
        Me.ccalcPictureBox.Location = New System.Drawing.Point(12, 37)
        Me.ccalcPictureBox.Name = "ccalcPictureBox"
        Me.ccalcPictureBox.Size = New System.Drawing.Size(111, 111)
        Me.ccalcPictureBox.TabIndex = 2
        Me.ccalcPictureBox.TabStop = False
        '
        'okButton
        '
        Me.okButton.Location = New System.Drawing.Point(324, 261)
        Me.okButton.Name = "okButton"
        Me.okButton.Size = New System.Drawing.Size(84, 32)
        Me.okButton.TabIndex = 3
        Me.okButton.Text = "OK"
        Me.okButton.UseVisualStyleBackColor = True
        '
        'frm_About
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(448, 305)
        Me.Controls.Add(Me.okButton)
        Me.Controls.Add(Me.ccalcPictureBox)
        Me.Controls.Add(Me.deptidTextBox)
        Me.Controls.Add(Me.copywriteTextBox)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "frm_About"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "About the Program"
        CType(Me.ccalcPictureBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents copywriteTextBox As System.Windows.Forms.TextBox
    Friend WithEvents deptidTextBox As System.Windows.Forms.TextBox
    Friend WithEvents ccalcPictureBox As System.Windows.Forms.PictureBox
    Friend WithEvents okButton As System.Windows.Forms.Button
End Class
