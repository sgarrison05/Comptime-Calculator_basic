<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_Main
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
        Me.clearButton = New System.Windows.Forms.Button()
        Me.exitButton = New System.Windows.Forms.Button()
        Me.calcButton = New System.Windows.Forms.Button()
        Me.optionsGroupBox = New System.Windows.Forms.GroupBox()
        Me.btn_ReconcileData = New System.Windows.Forms.Button()
        Me.btn_Email = New System.Windows.Forms.Button()
        Me.applyButton = New System.Windows.Forms.Button()
        Me.comptimeMenuStrip = New System.Windows.Forms.MenuStrip()
        Me.fileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ComptimerunToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.BankToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EditToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ClearToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CommandsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CalculateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ApplyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReadMeFileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.calcearnedTextBox = New System.Windows.Forms.TextBox()
        Me.lbl_CalcResultID = New System.Windows.Forms.Label()
        Me.accruedDateTimePicker = New System.Windows.Forms.DateTimePicker()
        Me.accruedRadioButton = New System.Windows.Forms.RadioButton()
        Me.spentRadioButton = New System.Windows.Forms.RadioButton()
        Me.actionGroupBox = New System.Windows.Forms.GroupBox()
        Me.actiondateLabel = New System.Windows.Forms.Label()
        Me.casereasonLabel = New System.Windows.Forms.Label()
        Me.earnedidLabel = New System.Windows.Forms.Label()
        Me.takenidLabel = New System.Windows.Forms.Label()
        Me.newbalidLabel = New System.Windows.Forms.Label()
        Me.newbalLabel = New System.Windows.Forms.Label()
        Me.prevbalLabel = New System.Windows.Forms.Label()
        Me.prevbalidLabel = New System.Windows.Forms.Label()
        Me.lineLabel = New System.Windows.Forms.Label()
        Me.hrs1Label = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.earnedTextBox = New System.Windows.Forms.TextBox()
        Me.takenTextBox = New System.Windows.Forms.TextBox()
        Me.caseComboBox = New System.Windows.Forms.ComboBox()
        Me.hrs2Label = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.optionsGroupBox.SuspendLayout()
        Me.comptimeMenuStrip.SuspendLayout()
        Me.actionGroupBox.SuspendLayout()
        Me.SuspendLayout()
        '
        'clearButton
        '
        Me.clearButton.BackColor = System.Drawing.SystemColors.ControlLight
        Me.clearButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.clearButton.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.clearButton.Location = New System.Drawing.Point(20, 19)
        Me.clearButton.Name = "clearButton"
        Me.clearButton.Size = New System.Drawing.Size(75, 32)
        Me.clearButton.TabIndex = 0
        Me.clearButton.Text = "Clear Form"
        Me.clearButton.UseVisualStyleBackColor = False
        '
        'exitButton
        '
        Me.exitButton.BackColor = System.Drawing.SystemColors.ControlLight
        Me.exitButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.exitButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.exitButton.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.exitButton.Location = New System.Drawing.Point(20, 252)
        Me.exitButton.Name = "exitButton"
        Me.exitButton.Size = New System.Drawing.Size(75, 32)
        Me.exitButton.TabIndex = 5
        Me.exitButton.Text = "Exit Application"
        Me.exitButton.UseVisualStyleBackColor = False
        '
        'calcButton
        '
        Me.calcButton.BackColor = System.Drawing.SystemColors.ControlLight
        Me.calcButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.calcButton.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.calcButton.Location = New System.Drawing.Point(20, 64)
        Me.calcButton.Name = "calcButton"
        Me.calcButton.Size = New System.Drawing.Size(75, 32)
        Me.calcButton.TabIndex = 1
        Me.calcButton.Text = "Calculation Preview"
        Me.calcButton.UseVisualStyleBackColor = False
        '
        'optionsGroupBox
        '
        Me.optionsGroupBox.BackColor = System.Drawing.Color.FromArgb(CType(CType(55, Byte), Integer), CType(CType(21, Byte), Integer), CType(CType(83, Byte), Integer))
        Me.optionsGroupBox.Controls.Add(Me.btn_ReconcileData)
        Me.optionsGroupBox.Controls.Add(Me.btn_Email)
        Me.optionsGroupBox.Controls.Add(Me.applyButton)
        Me.optionsGroupBox.Controls.Add(Me.clearButton)
        Me.optionsGroupBox.Controls.Add(Me.calcButton)
        Me.optionsGroupBox.Controls.Add(Me.exitButton)
        Me.optionsGroupBox.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.optionsGroupBox.Location = New System.Drawing.Point(477, 53)
        Me.optionsGroupBox.Name = "optionsGroupBox"
        Me.optionsGroupBox.Size = New System.Drawing.Size(117, 305)
        Me.optionsGroupBox.TabIndex = 5
        Me.optionsGroupBox.TabStop = False
        Me.optionsGroupBox.Text = "Options"
        '
        'btn_ReconcileData
        '
        Me.btn_ReconcileData.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btn_ReconcileData.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btn_ReconcileData.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_ReconcileData.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.btn_ReconcileData.Location = New System.Drawing.Point(20, 156)
        Me.btn_ReconcileData.Name = "btn_ReconcileData"
        Me.btn_ReconcileData.Size = New System.Drawing.Size(75, 32)
        Me.btn_ReconcileData.TabIndex = 3
        Me.btn_ReconcileData.Text = "Reconcile Data"
        Me.btn_ReconcileData.UseVisualStyleBackColor = False
        '
        'btn_Email
        '
        Me.btn_Email.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btn_Email.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btn_Email.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_Email.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.btn_Email.Location = New System.Drawing.Point(20, 204)
        Me.btn_Email.Name = "btn_Email"
        Me.btn_Email.Size = New System.Drawing.Size(75, 32)
        Me.btn_Email.TabIndex = 4
        Me.btn_Email.Text = "Email Data"
        Me.btn_Email.UseVisualStyleBackColor = False
        '
        'applyButton
        '
        Me.applyButton.BackColor = System.Drawing.SystemColors.ControlLight
        Me.applyButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.applyButton.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.applyButton.Location = New System.Drawing.Point(20, 108)
        Me.applyButton.Name = "applyButton"
        Me.applyButton.Size = New System.Drawing.Size(75, 32)
        Me.applyButton.TabIndex = 2
        Me.applyButton.Text = "Apply to Bank"
        Me.applyButton.UseVisualStyleBackColor = False
        '
        'comptimeMenuStrip
        '
        Me.comptimeMenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.fileToolStripMenuItem, Me.EditToolStripMenuItem, Me.HelpToolStripMenuItem, Me.HelpToolStripMenuItem1})
        Me.comptimeMenuStrip.Location = New System.Drawing.Point(0, 0)
        Me.comptimeMenuStrip.Name = "comptimeMenuStrip"
        Me.comptimeMenuStrip.Size = New System.Drawing.Size(609, 24)
        Me.comptimeMenuStrip.TabIndex = 6
        Me.comptimeMenuStrip.Text = "ComptimeMenuStrip"
        '
        'fileToolStripMenuItem
        '
        Me.fileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OpenToolStripMenuItem, Me.ToolStripSeparator1, Me.ExitToolStripMenuItem})
        Me.fileToolStripMenuItem.Name = "fileToolStripMenuItem"
        Me.fileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.fileToolStripMenuItem.Text = "&File"
        '
        'OpenToolStripMenuItem
        '
        Me.OpenToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ComptimerunToolStripMenuItem, Me.BankToolStripMenuItem})
        Me.OpenToolStripMenuItem.Name = "OpenToolStripMenuItem"
        Me.OpenToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.O), System.Windows.Forms.Keys)
        Me.OpenToolStripMenuItem.Size = New System.Drawing.Size(148, 22)
        Me.OpenToolStripMenuItem.Text = "Open.."
        '
        'ComptimerunToolStripMenuItem
        '
        Me.ComptimerunToolStripMenuItem.Name = "ComptimerunToolStripMenuItem"
        Me.ComptimerunToolStripMenuItem.Size = New System.Drawing.Size(146, 22)
        Me.ComptimerunToolStripMenuItem.Text = "Activity Sheet"
        '
        'BankToolStripMenuItem
        '
        Me.BankToolStripMenuItem.Name = "BankToolStripMenuItem"
        Me.BankToolStripMenuItem.Size = New System.Drawing.Size(146, 22)
        Me.BankToolStripMenuItem.Text = "Bank"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(145, 6)
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.E), System.Windows.Forms.Keys)
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(148, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'EditToolStripMenuItem
        '
        Me.EditToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ClearToolStripMenuItem})
        Me.EditToolStripMenuItem.Name = "EditToolStripMenuItem"
        Me.EditToolStripMenuItem.Size = New System.Drawing.Size(39, 20)
        Me.EditToolStripMenuItem.Text = "Edit"
        '
        'ClearToolStripMenuItem
        '
        Me.ClearToolStripMenuItem.Name = "ClearToolStripMenuItem"
        Me.ClearToolStripMenuItem.Size = New System.Drawing.Size(101, 22)
        Me.ClearToolStripMenuItem.Text = "Clear"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CommandsToolStripMenuItem})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem.Text = "&View"
        '
        'CommandsToolStripMenuItem
        '
        Me.CommandsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CalculateToolStripMenuItem, Me.ApplyToolStripMenuItem})
        Me.CommandsToolStripMenuItem.Name = "CommandsToolStripMenuItem"
        Me.CommandsToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me.CommandsToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.CommandsToolStripMenuItem.Text = "Commands.."
        '
        'CalculateToolStripMenuItem
        '
        Me.CalculateToolStripMenuItem.Name = "CalculateToolStripMenuItem"
        Me.CalculateToolStripMenuItem.Size = New System.Drawing.Size(108, 22)
        Me.CalculateToolStripMenuItem.Text = "Calc"
        '
        'ApplyToolStripMenuItem
        '
        Me.ApplyToolStripMenuItem.Name = "ApplyToolStripMenuItem"
        Me.ApplyToolStripMenuItem.Size = New System.Drawing.Size(108, 22)
        Me.ApplyToolStripMenuItem.Text = "Apply "
        '
        'HelpToolStripMenuItem1
        '
        Me.HelpToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AboutToolStripMenuItem, Me.ReadMeFileToolStripMenuItem})
        Me.HelpToolStripMenuItem1.Name = "HelpToolStripMenuItem1"
        Me.HelpToolStripMenuItem1.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem1.Text = "&Help"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(143, 22)
        Me.AboutToolStripMenuItem.Text = "About"
        '
        'ReadMeFileToolStripMenuItem
        '
        Me.ReadMeFileToolStripMenuItem.Name = "ReadMeFileToolStripMenuItem"
        Me.ReadMeFileToolStripMenuItem.Size = New System.Drawing.Size(143, 22)
        Me.ReadMeFileToolStripMenuItem.Text = "Read-Me File"
        '
        'calcearnedTextBox
        '
        Me.calcearnedTextBox.BackColor = System.Drawing.Color.Lavender
        Me.calcearnedTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.calcearnedTextBox.Location = New System.Drawing.Point(12, 366)
        Me.calcearnedTextBox.Multiline = True
        Me.calcearnedTextBox.Name = "calcearnedTextBox"
        Me.calcearnedTextBox.Size = New System.Drawing.Size(582, 131)
        Me.calcearnedTextBox.TabIndex = 22
        '
        'lbl_CalcResultID
        '
        Me.lbl_CalcResultID.AutoSize = True
        Me.lbl_CalcResultID.Location = New System.Drawing.Point(16, 350)
        Me.lbl_CalcResultID.Name = "lbl_CalcResultID"
        Me.lbl_CalcResultID.Size = New System.Drawing.Size(103, 13)
        Me.lbl_CalcResultID.TabIndex = 23
        Me.lbl_CalcResultID.Text = "Calculation Preview:"
        '
        'accruedDateTimePicker
        '
        Me.accruedDateTimePicker.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.accruedDateTimePicker.CustomFormat = "MM/dd/yyyy"
        Me.accruedDateTimePicker.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.accruedDateTimePicker.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.accruedDateTimePicker.Location = New System.Drawing.Point(16, 69)
        Me.accruedDateTimePicker.Name = "accruedDateTimePicker"
        Me.accruedDateTimePicker.Size = New System.Drawing.Size(120, 21)
        Me.accruedDateTimePicker.TabIndex = 0
        '
        'accruedRadioButton
        '
        Me.accruedRadioButton.AutoSize = True
        Me.accruedRadioButton.Location = New System.Drawing.Point(8, 21)
        Me.accruedRadioButton.Name = "accruedRadioButton"
        Me.accruedRadioButton.Size = New System.Drawing.Size(59, 17)
        Me.accruedRadioButton.TabIndex = 0
        Me.accruedRadioButton.TabStop = True
        Me.accruedRadioButton.Text = "Earned"
        Me.accruedRadioButton.UseVisualStyleBackColor = True
        '
        'spentRadioButton
        '
        Me.spentRadioButton.AutoSize = True
        Me.spentRadioButton.Location = New System.Drawing.Point(8, 65)
        Me.spentRadioButton.Name = "spentRadioButton"
        Me.spentRadioButton.Size = New System.Drawing.Size(56, 17)
        Me.spentRadioButton.TabIndex = 1
        Me.spentRadioButton.TabStop = True
        Me.spentRadioButton.Text = "Taken"
        Me.spentRadioButton.UseVisualStyleBackColor = True
        '
        'actionGroupBox
        '
        Me.actionGroupBox.BackColor = System.Drawing.Color.FromArgb(CType(CType(55, Byte), Integer), CType(CType(21, Byte), Integer), CType(CType(83, Byte), Integer))
        Me.actionGroupBox.Controls.Add(Me.spentRadioButton)
        Me.actionGroupBox.Controls.Add(Me.accruedRadioButton)
        Me.actionGroupBox.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.actionGroupBox.Location = New System.Drawing.Point(16, 117)
        Me.actionGroupBox.Name = "actionGroupBox"
        Me.actionGroupBox.Size = New System.Drawing.Size(120, 94)
        Me.actionGroupBox.TabIndex = 2
        Me.actionGroupBox.TabStop = False
        Me.actionGroupBox.Text = "Accrued / Spent"
        '
        'actiondateLabel
        '
        Me.actiondateLabel.AutoSize = True
        Me.actiondateLabel.Location = New System.Drawing.Point(16, 53)
        Me.actiondateLabel.Name = "actiondateLabel"
        Me.actiondateLabel.Size = New System.Drawing.Size(66, 13)
        Me.actiondateLabel.TabIndex = 0
        Me.actiondateLabel.Text = "Action Date:"
        '
        'casereasonLabel
        '
        Me.casereasonLabel.AutoSize = True
        Me.casereasonLabel.Location = New System.Drawing.Point(180, 53)
        Me.casereasonLabel.Name = "casereasonLabel"
        Me.casereasonLabel.Size = New System.Drawing.Size(118, 13)
        Me.casereasonLabel.TabIndex = 2
        Me.casereasonLabel.Text = "Case/Reason for Entry:"
        '
        'earnedidLabel
        '
        Me.earnedidLabel.AutoSize = True
        Me.earnedidLabel.Location = New System.Drawing.Point(180, 105)
        Me.earnedidLabel.Name = "earnedidLabel"
        Me.earnedidLabel.Size = New System.Drawing.Size(70, 13)
        Me.earnedidLabel.TabIndex = 5
        Me.earnedidLabel.Text = "Time Earned:"
        '
        'takenidLabel
        '
        Me.takenidLabel.AutoSize = True
        Me.takenidLabel.Location = New System.Drawing.Point(180, 173)
        Me.takenidLabel.Name = "takenidLabel"
        Me.takenidLabel.Size = New System.Drawing.Size(67, 13)
        Me.takenidLabel.TabIndex = 9
        Me.takenidLabel.Text = "Time Taken:"
        '
        'newbalidLabel
        '
        Me.newbalidLabel.AutoSize = True
        Me.newbalidLabel.Location = New System.Drawing.Point(206, 234)
        Me.newbalidLabel.Name = "newbalidLabel"
        Me.newbalidLabel.Size = New System.Drawing.Size(75, 13)
        Me.newbalidLabel.TabIndex = 14
        Me.newbalidLabel.Text = "Daily Balance:"
        '
        'newbalLabel
        '
        Me.newbalLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.newbalLabel.Location = New System.Drawing.Point(213, 262)
        Me.newbalLabel.Name = "newbalLabel"
        Me.newbalLabel.Size = New System.Drawing.Size(100, 23)
        Me.newbalLabel.TabIndex = 15
        '
        'prevbalLabel
        '
        Me.prevbalLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.prevbalLabel.Location = New System.Drawing.Point(21, 262)
        Me.prevbalLabel.Name = "prevbalLabel"
        Me.prevbalLabel.Size = New System.Drawing.Size(100, 23)
        Me.prevbalLabel.TabIndex = 13
        '
        'prevbalidLabel
        '
        Me.prevbalidLabel.AutoSize = True
        Me.prevbalidLabel.Location = New System.Drawing.Point(16, 234)
        Me.prevbalidLabel.Name = "prevbalidLabel"
        Me.prevbalidLabel.Size = New System.Drawing.Size(77, 13)
        Me.prevbalidLabel.TabIndex = 12
        Me.prevbalidLabel.Text = "Bank Balance:"
        '
        'lineLabel
        '
        Me.lineLabel.BackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.lineLabel.Location = New System.Drawing.Point(11, 227)
        Me.lineLabel.Name = "lineLabel"
        Me.lineLabel.Size = New System.Drawing.Size(326, 1)
        Me.lineLabel.TabIndex = 16
        '
        'hrs1Label
        '
        Me.hrs1Label.AutoSize = True
        Me.hrs1Label.Location = New System.Drawing.Point(237, 197)
        Me.hrs1Label.Name = "hrs1Label"
        Me.hrs1Label.Size = New System.Drawing.Size(24, 13)
        Me.hrs1Label.TabIndex = 11
        Me.hrs1Label.Text = "hrs "
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(164, 228)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(1, 67)
        Me.Label1.TabIndex = 17
        '
        'earnedTextBox
        '
        Me.earnedTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.earnedTextBox.Location = New System.Drawing.Point(183, 121)
        Me.earnedTextBox.MaxLength = 5
        Me.earnedTextBox.Name = "earnedTextBox"
        Me.earnedTextBox.Size = New System.Drawing.Size(47, 21)
        Me.earnedTextBox.TabIndex = 3
        '
        'takenTextBox
        '
        Me.takenTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.takenTextBox.Location = New System.Drawing.Point(183, 189)
        Me.takenTextBox.MaxLength = 5
        Me.takenTextBox.Name = "takenTextBox"
        Me.takenTextBox.Size = New System.Drawing.Size(47, 21)
        Me.takenTextBox.TabIndex = 4
        Me.takenTextBox.Tag = ""
        '
        'caseComboBox
        '
        Me.caseComboBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.caseComboBox.FormattingEnabled = True
        Me.caseComboBox.Location = New System.Drawing.Point(183, 69)
        Me.caseComboBox.Name = "caseComboBox"
        Me.caseComboBox.Size = New System.Drawing.Size(108, 23)
        Me.caseComboBox.TabIndex = 1
        '
        'hrs2Label
        '
        Me.hrs2Label.AutoSize = True
        Me.hrs2Label.Location = New System.Drawing.Point(237, 126)
        Me.hrs2Label.Name = "hrs2Label"
        Me.hrs2Label.Size = New System.Drawing.Size(24, 13)
        Me.hrs2Label.TabIndex = 19
        Me.hrs2Label.Text = "hrs "
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.Label3.Location = New System.Drawing.Point(11, 294)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(326, 1)
        Me.Label3.TabIndex = 20
        '
        'frm_Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.exitButton
        Me.ClientSize = New System.Drawing.Size(609, 509)
        Me.Controls.Add(Me.lbl_CalcResultID)
        Me.Controls.Add(Me.calcearnedTextBox)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.hrs2Label)
        Me.Controls.Add(Me.caseComboBox)
        Me.Controls.Add(Me.takenTextBox)
        Me.Controls.Add(Me.earnedTextBox)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.hrs1Label)
        Me.Controls.Add(Me.lineLabel)
        Me.Controls.Add(Me.prevbalidLabel)
        Me.Controls.Add(Me.prevbalLabel)
        Me.Controls.Add(Me.newbalLabel)
        Me.Controls.Add(Me.newbalidLabel)
        Me.Controls.Add(Me.takenidLabel)
        Me.Controls.Add(Me.earnedidLabel)
        Me.Controls.Add(Me.casereasonLabel)
        Me.Controls.Add(Me.optionsGroupBox)
        Me.Controls.Add(Me.actiondateLabel)
        Me.Controls.Add(Me.actionGroupBox)
        Me.Controls.Add(Me.accruedDateTimePicker)
        Me.Controls.Add(Me.comptimeMenuStrip)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MainMenuStrip = Me.comptimeMenuStrip
        Me.MaximizeBox = False
        Me.Name = "frm_Main"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Comptime Calculator"
        Me.optionsGroupBox.ResumeLayout(False)
        Me.comptimeMenuStrip.ResumeLayout(False)
        Me.comptimeMenuStrip.PerformLayout()
        Me.actionGroupBox.ResumeLayout(False)
        Me.actionGroupBox.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents clearButton As System.Windows.Forms.Button
    Friend WithEvents exitButton As System.Windows.Forms.Button
    Friend WithEvents calcButton As System.Windows.Forms.Button
    Friend WithEvents optionsGroupBox As System.Windows.Forms.GroupBox
    Friend WithEvents applyButton As System.Windows.Forms.Button
    Friend WithEvents comptimeMenuStrip As System.Windows.Forms.MenuStrip
    Friend WithEvents fileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OpenToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents CommandsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CalculateToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ApplyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ComptimerunToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BankToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents calcearnedTextBox As System.Windows.Forms.TextBox
    Friend WithEvents ReadMeFileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btn_ReconcileData As System.Windows.Forms.Button
    Friend WithEvents btn_Email As System.Windows.Forms.Button
    Friend WithEvents lbl_CalcResultID As System.Windows.Forms.Label
    Friend WithEvents EditToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ClearToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents accruedDateTimePicker As System.Windows.Forms.DateTimePicker
    Friend WithEvents accruedRadioButton As System.Windows.Forms.RadioButton
    Friend WithEvents spentRadioButton As System.Windows.Forms.RadioButton
    Friend WithEvents actionGroupBox As System.Windows.Forms.GroupBox
    Friend WithEvents actiondateLabel As System.Windows.Forms.Label
    Friend WithEvents casereasonLabel As System.Windows.Forms.Label
    Friend WithEvents earnedidLabel As System.Windows.Forms.Label
    Friend WithEvents takenidLabel As System.Windows.Forms.Label
    Friend WithEvents newbalidLabel As System.Windows.Forms.Label
    Friend WithEvents newbalLabel As System.Windows.Forms.Label
    Friend WithEvents prevbalLabel As System.Windows.Forms.Label
    Friend WithEvents prevbalidLabel As System.Windows.Forms.Label
    Friend WithEvents lineLabel As System.Windows.Forms.Label
    Friend WithEvents hrs1Label As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents earnedTextBox As System.Windows.Forms.TextBox
    Friend WithEvents takenTextBox As System.Windows.Forms.TextBox
    Friend WithEvents caseComboBox As System.Windows.Forms.ComboBox
    Friend WithEvents hrs2Label As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label

End Class
