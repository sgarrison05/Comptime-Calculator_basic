<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_Reconcillation
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
        Me.txt_ReconPrev = New System.Windows.Forms.TextBox()
        Me.btn_Return = New System.Windows.Forms.Button()
        Me.btn_Preview = New System.Windows.Forms.Button()
        Me.btn_Clear = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txt_ReconPrev
        '
        Me.txt_ReconPrev.Location = New System.Drawing.Point(12, 128)
        Me.txt_ReconPrev.Multiline = True
        Me.txt_ReconPrev.Name = "txt_ReconPrev"
        Me.txt_ReconPrev.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txt_ReconPrev.Size = New System.Drawing.Size(673, 439)
        Me.txt_ReconPrev.TabIndex = 0
        '
        'btn_Return
        '
        Me.btn_Return.Location = New System.Drawing.Point(240, 12)
        Me.btn_Return.Name = "btn_Return"
        Me.btn_Return.Size = New System.Drawing.Size(75, 37)
        Me.btn_Return.TabIndex = 1
        Me.btn_Return.Text = "Rtn to Main"
        Me.btn_Return.UseVisualStyleBackColor = True
        '
        'btn_Preview
        '
        Me.btn_Preview.Location = New System.Drawing.Point(12, 12)
        Me.btn_Preview.Name = "btn_Preview"
        Me.btn_Preview.Size = New System.Drawing.Size(75, 37)
        Me.btn_Preview.TabIndex = 2
        Me.btn_Preview.Text = "Preview"
        Me.btn_Preview.UseVisualStyleBackColor = True
        '
        'btn_Clear
        '
        Me.btn_Clear.Location = New System.Drawing.Point(123, 12)
        Me.btn_Clear.Name = "btn_Clear"
        Me.btn_Clear.Size = New System.Drawing.Size(75, 37)
        Me.btn_Clear.TabIndex = 3
        Me.btn_Clear.Text = "Clear"
        Me.btn_Clear.UseVisualStyleBackColor = True
        '
        'frm_Reconcillation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(697, 579)
        Me.Controls.Add(Me.btn_Clear)
        Me.Controls.Add(Me.btn_Preview)
        Me.Controls.Add(Me.btn_Return)
        Me.Controls.Add(Me.txt_ReconPrev)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "frm_Reconcillation"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Reconcillation Form"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txt_ReconPrev As System.Windows.Forms.TextBox
    Friend WithEvents btn_Return As System.Windows.Forms.Button
    Friend WithEvents btn_Preview As System.Windows.Forms.Button
    Friend WithEvents btn_Clear As System.Windows.Forms.Button
End Class
