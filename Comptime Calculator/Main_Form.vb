'Title                  Comptime Calculator
'Purpose                To calculate comptime time earned or spent in a particular instance
'Created By             Shon Garrison, December 2008 as part of Final for LSC-O VB Class
'Updated Last           April 2026

'Update Notes:          Refactored based on a collaboration with Claude.ai

Option Explicit On

Public Class frm_Main

    Friend Const CDIRECTORY As String = "D:\Temp\Comptime"
    Friend Const CPATH As String = "D:\Temp\Comptime\comptimerun.txt"
    Friend Const TITLE As String = "Comptime Calculator"
    Private Const WARNING_HOURS As Decimal = 50D
    Private Const COMP_TIME_MULTIPLIER As Decimal = 1.5D

    Public userName As String
    Private _newBalance As Decimal
    Private _previousBalance As Decimal
    Private _lastEntryBalance As String
    Private _myentry As String

    Friend ReadOnly _heading As String =
        "Date Entered" & Strings.Space(7) &
        "CaseNo." & Strings.Space(13) &
        "Earned(+)" & Strings.Space(7) &
        "Taken(-)" & Strings.Space(9) &
        "Balance"

    Friend ReadOnly _columnDivider As String =
        "-------------" & Strings.Space(6) &
        "----------" & Strings.Space(10) &
        "----------" & Strings.Space(6) &
        "----------" & Strings.Space(7) &
        "----------"

    '--------------------------------------------  Events -------------------------------------------------------------------------
    Private Sub compcalcForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        InitializeControls()

        If My.Computer.FileSystem.FileExists(CPATH) Then
            LoadExistingBankFile()
        Else
            PromptToCreateBankFile()
        End If

    End Sub

    Private Sub InitializeControls()

        Me.applyButton.Enabled = False
        Me.AboutToolStripMenuItem.Enabled = False

        'Case/Resason ComboBox
        With Me.caseComboBox.Items
            .Add("[Enter One]")
            .Add("Sick")
            .Add("Personal")
            .Add("Dr. Appt")
            .Add("New Case")
            .Add("Transport")
            .Add("Det Visit")
            .Add("On-Call")
            .Add("Spec Group")
            .Add("Plcmt Visit")
            .Add("Training")
            .Add("Evaluation")
            .Add("Meeting")
        End With
        Me.caseComboBox.SelectedItem = "[Enter One]"

        Me.accruedRadioButton.Select()
        Me.accruedDateTimePicker.Focus()

    End Sub

    Private Sub LoadExistingBankFile()

        Try
            Dim fileText As String = My.Computer.FileSystem.ReadAllText(CPATH)
            Dim entryIndex As Integer = 0
            Dim newLineIndex As Integer = fileText.IndexOf(ControlChars.NewLine, entryIndex)

            Do Until newLineIndex = -1
                Dim entry As String = fileText.Substring(entryIndex, newLineIndex - entryIndex)

                ' Extract the running balance from any line that contains a date
                If entry.Contains("/") Then
                    _lastEntryBalance = Trim(Microsoft.VisualBasic.Right(entry, 7))
                End If

                ' Extract the username from the account header line
                If entry.Contains("Account") Then
                    userName = entry.Substring(31)
                End If

                entryIndex = newLineIndex + 2
                newLineIndex = fileText.IndexOf(ControlChars.NewLine, entryIndex)
            Loop

            Me.prevbalLabel.Text = If(_lastEntryBalance, "0.00")
            Me.Text = "Personal Comptime Calculator for " & userName
            Me.newbalLabel.Text = "0.00"
            Me.calcearnedTextBox.Text = "Ready"

            UpdateWarningDisplay()

        Catch ex As Exception
            MessageBox.Show("Error reading comptime file:" & Environment.NewLine & ex.Message,
                            TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub PromptToCreateBankFile()

        Dim createResult As DialogResult =
            MessageBox.Show("The current comptime balance file does not exist. " &
                            "This is your comptime bank. Would you like to create it?",
                            TITLE, MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If createResult = DialogResult.No Then
            Me.Close()
            Return
        End If

        Dim hasBalance As DialogResult =
            MessageBox.Show("Do you have a current balance to enter?",
                            TITLE, MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If hasBalance = DialogResult.Yes Then

            ' Keep prompting until a valid numeric balance is entered
            Dim balanceInput As String = String.Empty
            Dim formattedValue As Decimal
            Do
                balanceInput = InputBox("Please enter current balance or click 'Ok' to start at zero.",
                                        TITLE, "0.00")
                If Not IsNumeric(balanceInput) Then
                    MessageBox.Show("Balance must be a number.", TITLE, MessageBoxButtons.OK)
                End If
            Loop Until IsNumeric(balanceInput)

            'conversion for two decimal places
            Decimal.TryParse(balanceInput, formattedValue)
            Dim output As String = formattedValue.ToString("N2")

            userName = InputBox("Please enter your name.", TITLE)
            Me.prevbalLabel.Text = output

        Else
            Me.prevbalLabel.Text = "0.00"
        End If

        Me.Text = "Personal Comptime Calculator for " & userName
        Me.newbalLabel.Text = "0.00"
        Me.calcearnedTextBox.Text = "Ready"
        UpdateWarningDisplay()

    End Sub

    Private Sub UpdateWarningDisplay()

        Dim currentBalance As Decimal
        If Decimal.TryParse(Me.prevbalLabel.Text, currentBalance) AndAlso
           currentBalance >= WARNING_HOURS Then
            Me.prevbalLabel.ForeColor = Color.Red
            Me.warningLbl.Show()
        Else
            Me.prevbalLabel.ForeColor = Color.Black
            Me.warningLbl.Hide()
        End If

    End Sub

    '----------------------------------- Custom Subroutines and Functions ---------------------------------------------------------

    Public Sub ApplyCalculations()
        'Saves current balance to txt file


        Dim applyResult As DialogResult =
                    MessageBox.Show("Do you wish to add the new balance to the bank?",
                                    TITLE, MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If applyResult = DialogResult.Yes Then

            CommitBalanceUpdate()
            WriteTransactionToFile()

            MessageBox.Show("Processing complete. The form will be cleared.",
                            TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            Dim continueResult As DialogResult =
                MessageBox.Show("Do you want to make another calculation?",
                                TITLE, MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If continueResult = DialogResult.No Then
                MessageBox.Show("No calculation will be made and the form will be reset. " &
                                "You may exit the program.",
                                TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If

        ' Reset the form regardless of whether the user saved or not
        CleanHouse()
        UpdateWarningDisplay()

    End Sub

    Public Sub CreatePlaceholderEntry()

        'Only used as a placeholder on first run if no prior transaction is completed

        Try
            _newBalance = Math.Round(Convert.ToDecimal(Me.newbalLabel.Text), 2)
            _previousBalance = Math.Round(Convert.ToDecimal(Me.prevbalLabel.Text) + _newBalance, 2)
            Me.prevbalLabel.Text = Convert.ToString(_previousBalance)

            If Not My.Computer.FileSystem.FileExists(CPATH) Then
                WriteFileHeader()
            End If

            Dim placeholderRow As String =
                Me.accruedDateTimePicker.Text & Strings.Space(9) &
                "Placeholder".PadRight(15, " "c) & Strings.Space(4) &
                "0.00".PadLeft(5, " "c) & Strings.Space(11) &
                "0.00".PadLeft(5, " "c) & Strings.Space(13) &
                Convert.ToString(_previousBalance).PadLeft(5, " "c) & ControlChars.NewLine

            My.Computer.FileSystem.WriteAllText(CPATH, placeholderRow, True)
            Separation()

        Catch ex As Exception
            MessageBox.Show("Error creating placeholder entry:" & Environment.NewLine & ex.Message,
                            TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Public Sub PreviewCalculations()

        'declare calculation variables
        Dim earned As Decimal
        Dim taken As Decimal
        Dim calcEarned As Decimal
        Dim previewBalance As Decimal

        'Determine if this time is accrued or taken
        If accruedRadioButton.Checked Then
            If takenTextBox.Text = String.Empty Then takenTextBox.Text = "0.00"

            If Not Decimal.TryParse(earnedTextBox.Text, earned) OrElse
            Not Decimal.TryParse(takenTextBox.Text, taken) Then
                MessageBox.Show("Hours entered must be numeric.", TITLE,
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.earnedTextBox.Focus()
                Return
            End If

            'If successful, make calculations
            calcEarned = Math.Round(earned * COMP_TIME_MULTIPLIER, 2)
            previewBalance = calcEarned + Convert.ToDecimal(prevbalLabel.Text)
            _newBalance = Math.Round(calcEarned - taken, 2)
            Me.newbalLabel.Text = Convert.ToString(_newBalance)

            Me.calcearnedTextBox.Text = "Total accrued time to enter on affidavit = " &
                calcEarned.ToString("N2") &
                " hours" &
                BuildPreviewText(previewBalance)

        ElseIf spentRadioButton.Checked Then

            If earnedTextBox.Text = String.Empty Then earnedTextBox.Text = "0.00"

            If Not Decimal.TryParse(Me.earnedTextBox.Text, earned) OrElse
               Not Decimal.TryParse(Me.takenTextBox.Text, taken) Then
                MessageBox.Show("Hours entered must be numeric.", TITLE,
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.takenTextBox.Focus()
                Return
            End If

            calcEarned = Math.Round(earned * COMP_TIME_MULTIPLIER, 2)
            _newBalance = Math.Round(calcEarned - taken, 2)
            previewBalance = _newBalance + Convert.ToDecimal(prevbalLabel.Text)
            newbalLabel.Text = Convert.ToString(_newBalance)
            calcearnedTextBox.Text = ""

            calcearnedTextBox.Text = "Total taken time to enter on affidavit = " &
               taken.ToString("N2") &
               " hours" &
               BuildPreviewText(previewBalance)
        End If

        Me.applyButton.Enabled = True
        Me.ApplyToolStripMenuItem.Enabled = True



    End Sub

    Public Sub CalcClear()

        'clears the form 
        Dim saveResult As DialogResult =
            MessageBox.Show("Do you wish to add the new balance to the bank?",
                            TITLE, MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If saveResult = DialogResult.Yes Then
            CommitBalanceUpdate()
            WriteTransactionToFile()
        End If

        CleanHouse()
        UpdateWarningDisplay()

    End Sub

    Private Sub exitApp()

        'Exits the Program

        Dim exitResult As DialogResult
        exitResult = MessageBox.Show("Are you sure that you are ready to exit?", TITLE,
                                    MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If exitResult = Windows.Forms.DialogResult.No Then
            CleanHouse()
            UpdateWarningDisplay()
            Return
        End If

        If Not My.Computer.FileSystem.FileExists(CPATH) Then
            CreatePlaceholderEntry()
        End If

        Me.Close()

    End Sub

    Private Sub Separation()

        My.Computer.FileSystem.WriteAllText(CPATH, "".PadLeft(85, "-") &
                                            ControlChars.NewLine, True)

    End Sub

    Private Sub CleanHouse()

        newbalLabel.Text = "0.00"
        calcearnedTextBox.Text = "Ready"
        caseComboBox.SelectedItem = "[Enter One]"
        earnedTextBox.Clear()
        takenTextBox.Clear()
        accruedRadioButton.Select()
        accruedDateTimePicker.Focus()
        Me.applyButton.Enabled = False
        Me.ApplyToolStripMenuItem.Enabled = False
        _newBalance = 0D

    End Sub

    Private Sub CommitBalanceUpdate()

        _newBalance = Math.Round(Convert.ToDecimal(Me.newbalLabel.Text), 2)
        _previousBalance = Math.Round(Convert.ToDecimal(Me.prevbalLabel.Text) + _newBalance, 2)
        Me.prevbalLabel.Text = Convert.ToString(_previousBalance)

    End Sub

    Private Sub WriteTransactionToFile()

        Try
            Dim currentDate As String = Me.accruedDateTimePicker.Text
            Dim caseNo As String = Me.caseComboBox.Text
            Dim earnedText As String = Me.earnedTextBox.Text
            Dim takenText As String = Me.takenTextBox.Text
            Dim balanceText As String = Convert.ToString(_previousBalance)

            If Not My.Computer.FileSystem.FileExists(CPATH) Then
                WriteFileHeader()
            End If

            Dim transactionRow As String =
                currentDate & Strings.Space(9) &
                caseNo.PadRight(15, " "c) & Strings.Space(4) &
                earnedText.PadLeft(5, " "c) & Strings.Space(11) &
                takenText.PadLeft(5, " "c) & Strings.Space(13) &
                balanceText.PadLeft(5, " "c) & ControlChars.NewLine

            My.Computer.FileSystem.WriteAllText(CPATH, transactionRow, True)
            Separation()

        Catch ex As Exception
            MessageBox.Show("Error writing to comptime file:" & Environment.NewLine & ex.Message,
                            TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub WriteFileHeader()

        If Not My.Computer.FileSystem.DirectoryExists(CDIRECTORY) Then
            My.Computer.FileSystem.CreateDirectory(CDIRECTORY)
        End If

        My.Computer.FileSystem.WriteAllText(CPATH,
            "Orange County Juvenile Probation Dept" & ControlChars.NewLine &
            "---------------------------------------" & ControlChars.NewLine &
            "Personal Comptime Account for: " & userName & ControlChars.NewLine &
            ControlChars.NewLine &
            _heading & ControlChars.NewLine &
            _columnDivider & ControlChars.NewLine, True)

    End Sub

    Private Function BuildPreviewText(previewBalance As Decimal) As String

        Return ControlChars.NewLine &
               "=".PadLeft(72, "=") & ControlChars.NewLine &
               "Preview of Entry to Activity Sheet:" & ControlChars.NewLine & ControlChars.NewLine &
               "Date Entered" & Strings.Space(14) &
               "CaseNo." & Strings.Space(14) &
               "Earned(+)" & Strings.Space(12) &
               "Taken(-)" & Strings.Space(16) &
               "Balance" & ControlChars.NewLine &
               "-----------------" & Strings.Space(13) &
               "----------" & Strings.Space(16) &
               "----------" & Strings.Space(17) &
               "----------" & Strings.Space(17) &
               "----------" & ControlChars.NewLine &
               accruedDateTimePicker.Text & Strings.Space(16) &
               caseComboBox.Text.PadRight(15, " ") & Strings.Space(7) &
               earnedTextBox.Text.PadLeft(5, " ") & Strings.Space(22) &
               takenTextBox.Text.PadLeft(5, " ") & Strings.Space(22) &
               Convert.ToString(previewBalance).PadLeft(5, " ")

    End Function
    '---------------------------------------Buttons and Click Events -------------------------------------------------------------

    Private Sub clearButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles clearButton.Click

        CalcClear()

    End Sub

    Private Sub calcButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles calcButton.Click

        PreviewCalculations()

    End Sub

    Private Sub exitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles exitButton.Click

        exitApp()

    End Sub

    Private Sub btn_ReconcileData_Click(sender As System.Object, e As System.EventArgs) Handles btn_ReconcileData.Click

        Me.Hide()
        frm_Reconcile.Show()

    End Sub

    Private Sub ClearToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearToolStripMenuItem.Click
        CalcClear()
    End Sub

    Private Sub earnedTextBox_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
        earnedTextBox.SelectAll()

    End Sub

    Private Sub earnedTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        newbalLabel.Text = ""
        calcearnedTextBox.Text = ""
        calcearnedTextBox.Text = "Ready"
        takenTextBox.Clear()
    End Sub

    Private Sub caseComboBox_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
        caseComboBox.SelectAll()

    End Sub

    Private Sub caseComboBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        newbalLabel.Text = ""
        calcearnedTextBox.Text = ""
        calcearnedTextBox.Text = "Ready"
        earnedTextBox.Clear()
        takenTextBox.Clear()
    End Sub

    Private Sub takenTextBox_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
        takenTextBox.SelectAll()

    End Sub

    Private Sub takenTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        newbalLabel.Text = ""
        calcearnedTextBox.Text = ""
        calcearnedTextBox.Text = "Ready"
        earnedTextBox.Clear()
    End Sub

    Private Sub spentRadioButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles spentRadioButton.Click
        earnedTextBox.Enabled = False
        takenTextBox.Enabled = True

    End Sub

    Private Sub accruedRadioButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles accruedRadioButton.Click
        takenTextBox.Enabled = False
        earnedTextBox.Enabled = True
    End Sub

    Private Sub applyButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles applyButton.Click

        ApplyCalculations()

    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click

        exitApp()

    End Sub

    Private Sub ComptimerunToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComptimerunToolStripMenuItem.Click
        'Opens your comptime activity sheet file

        Dim proc As New System.Diagnostics.Process
        proc.StartInfo.FileName = "notepad.exe"
        proc.StartInfo.Arguments = CPATH

        proc.Start()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        frm_About.ShowDialog()
    End Sub

    Private Sub ReadMeFileToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ReadMeFileToolStripMenuItem.Click

        'Opens your comptime readme file

        Dim proc As New System.Diagnostics.Process
        proc.StartInfo.FileName = "notepad.exe"
        proc.StartInfo.Arguments = "C:/Comptime/Readme.txt"

        proc.Start()
    End Sub

    Private Sub ClearFormToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        CalcClear()

    End Sub

    Private Sub CalculateToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CalculateToolStripMenuItem.Click

        PreviewCalculations()

    End Sub

    Private Sub ApplyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ApplyToolStripMenuItem.Click

        ApplyCalculations()

    End Sub

End Class