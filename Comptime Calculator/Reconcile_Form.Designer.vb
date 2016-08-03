<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_Reconcile
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_Reconcile))
        Me.lblPreview = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnClearPrev = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnReconcile = New System.Windows.Forms.Button()
        Me.txtYear = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnPreview = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.libxPreview = New System.Windows.Forms.ListBox()
        Me.lblCPreview = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.libxOrig = New System.Windows.Forms.ListBox()
        Me.lblPreview_Orig = New System.Windows.Forms.Label()
        Me.libxReconcile = New System.Windows.Forms.ListBox()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblPreview
        '
        Me.lblPreview.AutoSize = True
        Me.lblPreview.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPreview.Location = New System.Drawing.Point(508, 56)
        Me.lblPreview.Name = "lblPreview"
        Me.lblPreview.Size = New System.Drawing.Size(344, 13)
        Me.lblPreview.TabIndex = 1
        Me.lblPreview.Text = "Preview of Previous Year -> (Entries that will be taken out):"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnClearPrev)
        Me.GroupBox1.Controls.Add(Me.btnExit)
        Me.GroupBox1.Controls.Add(Me.btnReconcile)
        Me.GroupBox1.Location = New System.Drawing.Point(931, 67)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(108, 229)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Options"
        '
        'btnClearPrev
        '
        Me.btnClearPrev.Location = New System.Drawing.Point(17, 92)
        Me.btnClearPrev.Name = "btnClearPrev"
        Me.btnClearPrev.Size = New System.Drawing.Size(75, 41)
        Me.btnClearPrev.TabIndex = 3
        Me.btnClearPrev.Text = "Clear Preview"
        Me.btnClearPrev.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(17, 166)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(75, 41)
        Me.btnExit.TabIndex = 2
        Me.btnExit.Text = "Rtn to Main"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnReconcile
        '
        Me.btnReconcile.Location = New System.Drawing.Point(17, 19)
        Me.btnReconcile.Name = "btnReconcile"
        Me.btnReconcile.Size = New System.Drawing.Size(75, 41)
        Me.btnReconcile.TabIndex = 0
        Me.btnReconcile.Text = "Reconcile"
        Me.btnReconcile.UseVisualStyleBackColor = True
        '
        'txtYear
        '
        Me.txtYear.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtYear.Location = New System.Drawing.Point(116, 4)
        Me.txtYear.Name = "txtYear"
        Me.txtYear.Size = New System.Drawing.Size(95, 29)
        Me.txtYear.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 4)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(95, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Year to Reconcile:"
        '
        'btnPreview
        '
        Me.btnPreview.Image = CType(resources.GetObject("btnPreview.Image"), System.Drawing.Image)
        Me.btnPreview.Location = New System.Drawing.Point(431, 194)
        Me.btnPreview.Name = "btnPreview"
        Me.btnPreview.Size = New System.Drawing.Size(60, 50)
        Me.btnPreview.TabIndex = 7
        Me.btnPreview.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(15, 56)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(169, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Current Datafile - All Entries:"
        '
        'libxPreview
        '
        Me.libxPreview.FormattingEnabled = True
        Me.libxPreview.Location = New System.Drawing.Point(511, 72)
        Me.libxPreview.Name = "libxPreview"
        Me.libxPreview.Size = New System.Drawing.Size(400, 264)
        Me.libxPreview.TabIndex = 9
        '
        'lblCPreview
        '
        Me.lblCPreview.Location = New System.Drawing.Point(428, 159)
        Me.lblCPreview.Name = "lblCPreview"
        Me.lblCPreview.Size = New System.Drawing.Size(48, 32)
        Me.lblCPreview.TabIndex = 10
        Me.lblCPreview.Text = " Click to Preview:"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label5.Location = New System.Drawing.Point(1, 36)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(1050, 10)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Label5"
        '
        'libxOrig
        '
        Me.libxOrig.FormattingEnabled = True
        Me.libxOrig.Location = New System.Drawing.Point(511, 357)
        Me.libxOrig.Name = "libxOrig"
        Me.libxOrig.Size = New System.Drawing.Size(400, 290)
        Me.libxOrig.TabIndex = 12
        '
        'lblPreview_Orig
        '
        Me.lblPreview_Orig.AutoSize = True
        Me.lblPreview_Orig.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPreview_Orig.Location = New System.Drawing.Point(508, 339)
        Me.lblPreview_Orig.Name = "lblPreview_Orig"
        Me.lblPreview_Orig.Size = New System.Drawing.Size(328, 13)
        Me.lblPreview_Orig.TabIndex = 13
        Me.lblPreview_Orig.Text = "Preview of Current Year -> (What will remain in Datafile):"
        '
        'libxReconcile
        '
        Me.libxReconcile.FormattingEnabled = True
        Me.libxReconcile.Location = New System.Drawing.Point(18, 73)
        Me.libxReconcile.Name = "libxReconcile"
        Me.libxReconcile.Size = New System.Drawing.Size(393, 576)
        Me.libxReconcile.TabIndex = 14
        '
        'frm_Reconcile
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1051, 667)
        Me.Controls.Add(Me.libxReconcile)
        Me.Controls.Add(Me.lblPreview_Orig)
        Me.Controls.Add(Me.libxOrig)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.lblCPreview)
        Me.Controls.Add(Me.libxPreview)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnPreview)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtYear)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.lblPreview)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "frm_Reconcile"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Reconcile Current Year"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblPreview As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents btnReconcile As System.Windows.Forms.Button
    Friend WithEvents txtYear As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnPreview As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnReload As System.Windows.Forms.Button
    Friend WithEvents libxPreview As System.Windows.Forms.ListBox
    Friend WithEvents lblCPreview As System.Windows.Forms.Label
    Friend WithEvents btnClearPrev As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents libxOrig As System.Windows.Forms.ListBox
    Friend WithEvents lblPreview_Orig As System.Windows.Forms.Label
    Friend WithEvents libxReconcile As System.Windows.Forms.ListBox
End Class
