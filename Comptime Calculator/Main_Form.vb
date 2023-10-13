'Title                  Comptime Calculator
'Purpose                To calculate comptime time earned or spent
'                       in a particular instance
'Created By             Shon Garrison, December 2008
'Updated Last           September 28, 2021

'Update Notes:          Elemenated the bank file

Option Explicit On

Public Class frm_Main

    Private Const cdirectory As String = "C:\Comptime"
    Private Const cpath As String = "C:\Comptime\comptimerun.txt"
    Private title As String = "Comptime Calculator"
    Public user As String
    Private newbalance As Decimal
    Private previous As Decimal
    Private myentry As String
    Private heading As String = "Date Entered" & Strings.Space(10) &
                                "CaseNo." & Strings.Space(13) &
                                "Earned(+)" & Strings.Space(10) &
                                "Taken(-)" & Strings.Space(9) &
                                "Balance"

    '--------------------------------------------  Events -------------------------------------------------------------------------
    Private Sub compcalcForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim my_decision As DialogResult

        Me.calcearnedTextBox.ReadOnly = True

        'disable apply calc button till preview is seen
        Me.applyButton.Enabled = False
        Me.ApplyToolStripMenuItem.Enabled = False

        Me.caseComboBox.Items.Add("[Enter One]")
        Me.caseComboBox.Items.Add("Sick")
        Me.caseComboBox.Items.Add("Personal")
        Me.caseComboBox.Items.Add("Dr. Appt")
        Me.caseComboBox.Items.Add("New Case")
        Me.caseComboBox.Items.Add("Transport")
        Me.caseComboBox.Items.Add("Det Visit")
        Me.caseComboBox.Items.Add("On-Call")
        Me.caseComboBox.Items.Add("Spec Group")
        Me.caseComboBox.Items.Add("Plcmt Visit")
        Me.caseComboBox.Items.Add("Training")
        Me.caseComboBox.Items.Add("Evaluation")
        Me.caseComboBox.Items.Add("Meeting")

        Me.caseComboBox.SelectedItem = "[Enter One]"

        'sets up a default selection for the radio buttons
        accruedRadioButton.Select()

        'Verify txt file exists else ask if the user wants to create it.
        If My.Computer.FileSystem.FileExists(cpath) Then

            'pulls user variable from text file
            'declare block variables
            Dim readtxt As String
            Dim entry As String
            Dim newLineIndex As Integer = 0
            Dim entryIndex As Integer = 0
            Dim entryuser As String

            'checks for existing comptimerun.txt
            'if it exists, it pulls it and stores it
            If My.Computer.FileSystem.FileExists(cpath) Then

                readtxt = My.Computer.FileSystem.ReadAllText(cpath)

                'primer for first read of readtxt
                newLineIndex = readtxt.IndexOf(ControlChars.NewLine, entryIndex)

                Do Until newLineIndex = -1

                    'get each line
                    entry = readtxt.Substring(entryIndex, newLineIndex - entryIndex)

                    'finds line  with username that you are searching for
                    entryuser = entry.Contains("Account")

                    'if line is found with a date, add to myentry variable to find bank balance at the end
                    If entry.Contains("/") Then
                        myentry = Trim(Microsoft.VisualBasic.Right(entry, 7))
                    End If

                    'Retrieve Current Bank balance
                    prevbalLabel.Text = myentry

                    'if user is found, it adds it to preview
                    If entryuser = True Then
                        user = entry.Substring(31)
                    End If

                    'if not found updates entryindex with next line
                    entryIndex = newLineIndex + 2
                    newLineIndex = readtxt.IndexOf(ControlChars.NewLine, entryIndex)

                Loop

                Me.Text = "Personal Comptime Calculator for " & user
                newbalLabel.Text = "0.00"
                calcearnedTextBox.Text = "Ready"

            End If
        Else 'sets up the comptime back in the specified path
            my_decision = MessageBox.Show _
            ("The current comptime balance file does not exist.  This is your comptime bank, would you like to create it?",
            "Comptime Calculator", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)

            If my_decision = DialogResult.Yes Then
                Dim init_bal_decision As DialogResult
                init_bal_decision = MessageBox.Show("Do you have an current balance to enter?", title, MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)

                If init_bal_decision = Windows.Forms.DialogResult.Yes Then
                    Do Until IsNumeric(prevbalLabel.Text) Or prevbalLabel.Text <> String.Empty
                        prevbalLabel.Text = InputBox("Please enter current balance or click 'Ok' to go to calculator.", title, "0.00")
                        If Not IsNumeric(prevbalLabel.Text) Then
                            MessageBox.Show("Number must be numeric.", title, MessageBoxButtons.OK)
                        End If
                    Loop

                    user = InputBox("Please Enter Your name", title, )

                    'Quick Conversion for two decimal places for label
                    Dim myConvert As Decimal = CDec(prevbalLabel.Text)
                    prevbalLabel.Text = myConvert.ToString("N2")
                    previous = CDec(prevbalLabel.Text)

                    Me.Show()
                    Me.Text = "Personal Comptime Calculator for " & user
                    newbalLabel.Text = "0.00"
                    calcearnedTextBox.Text = "Ready"

                ElseIf init_bal_decision = Windows.Forms.DialogResult.No Then
                    Me.Show()
                    prevbalLabel.Text = "0.00"
                    newbalLabel.Text = "0.00"
                End If
            Else : my_decision = DialogResult.No
                Me.Close()
            End If
        End If

    End Sub

    '----------------------------------- Custom Subroutines and Functions ---------------------------------------------------------

    Public Sub ApplyCalculations()
        'Saves current balance to txt file

        Dim my_apply As DialogResult
        Dim my_another As DialogResult

        'It will ask to you append file.
        my_apply = MessageBox.Show("Do you wish to add to new balance to bank?", title, MessageBoxButtons.YesNo,
        MessageBoxIcon.Question)

        If my_apply = Windows.Forms.DialogResult.Yes Then

            'make calculations
            newbalance = Convert.ToDecimal(newbalLabel.Text)
            newbalance = Math.Round(newbalance, 2)
            previous = Convert.ToDecimal(prevbalLabel.Text)
            previous += newbalance
            previous = Math.Round(previous, 2)
            prevbalLabel.Text = Convert.ToString(previous)

            'Declare text writing variables
            Dim line As String
            Dim caseno As String
            Dim curdate As String

            'Convert the data from the Previous Balance Label and store it in line variable
            line = Convert.ToString(previous)
            caseno = caseComboBox.Text
            curdate = accruedDateTimePicker.Text

            'If Comptime file exists, the prog writes current balance text file
            If My.Computer.FileSystem.FileExists(cpath) Then
                My.Computer.FileSystem.WriteAllText(cpath, curdate & Strings.Space(12) &
                                                    caseno.PadRight(15, " ") & Strings.Space(6) &
                                                    earnedTextBox.Text.PadLeft(5, " ") & Strings.Space(13) &
                                                    takenTextBox.Text.PadLeft(5, " ") & Strings.Space(13) &
                                                    Convert.ToString(previous).PadLeft(5, " ") &
                                                    ControlChars.NewLine, True)

                Separation()

                MessageBox.Show("Processing complete. The form will be cleared.",
                                title, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.Show()

                CleanHouse()

                'If Comptime file does not exists, the program creates it and
                'writes current balance text file
            Else : My.Computer.FileSystem.CreateDirectory(cdirectory)
                My.Computer.FileSystem.WriteAllText(cpath,
                                                    "Orange County Juvenile Probation Dept" & ControlChars.NewLine &
                                                    "---------------------------------------" & ControlChars.NewLine &
                                                    "Personal Comptime Account for: " & user & ControlChars.NewLine &
                                                    ControlChars.NewLine &
                                                    heading & ControlChars.NewLine &
                                                    "------------" & Strings.Space(10) &
                                                    "--------------" & Strings.Space(5) &
                                                    "---------" & Strings.Space(10) &
                                                    "--------" & Strings.Space(9) &
                                                    "-------" & ControlChars.NewLine, True)

                My.Computer.FileSystem.WriteAllText(cpath, curdate & Strings.Space(12) &
                                                    caseno.PadRight(15, " ") & Strings.Space(6) &
                                                    earnedTextBox.Text.PadLeft(5, " ") & Strings.Space(13) &
                                                    takenTextBox.Text.PadLeft(5, " ") & Strings.Space(13) &
                                                    Convert.ToString(previous).PadLeft(5, " ") &
                                                    ControlChars.NewLine, True)

                Separation()

                MessageBox.Show("Processing complete. The form will be cleared.",
                                title, MessageBoxButtons.OK, MessageBoxIcon.Information)

                Me.Show()

                CleanHouse()

            End If

            'If user does not want to make calculation, the program
            'will ask if the user wants to return to the program for another calculation.
            'If the user does not, then the program will direct user to exit.
        Else : my_apply = Windows.Forms.DialogResult.No
            my_another = MessageBox.Show("Do you want to make another calculation?", title,
        MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If my_another = Windows.Forms.DialogResult.Yes Then
                Me.Show()

                CleanHouse()

            Else : my_another = Windows.Forms.DialogResult.No
                MessageBox.Show("No calcuation will be made and the form will be reset. You may exit the program.", title,
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.Show()

                CleanHouse()

            End If

        End If

    End Sub

    Public Sub CreateMyPaths()

        'Only used as a placeholder on first run if no prior transaction is completed

        'Declare text writing variables
        Dim line As String
        Dim caseno As String = "Placeholder"
        Dim curdate As String = accruedDateTimePicker.Text

        'make calculations
        newbalance = Convert.ToDecimal(newbalLabel.Text)
        newbalance = Math.Round(newbalance, 2)
        previous = Convert.ToDecimal(prevbalLabel.Text)
        previous += newbalance
        previous = Math.Round(previous, 2)
        prevbalLabel.Text = Convert.ToString(previous)

        'Convert the data from the Previous Balance Label and store it in line variable
        line = Convert.ToString(previous)

        My.Computer.FileSystem.CreateDirectory(cdirectory)
        My.Computer.FileSystem.WriteAllText(cpath,
                                            "Orange County Juvenile Probation Dept" & ControlChars.NewLine &
                                            "---------------------------------------" & ControlChars.NewLine &
                                            "Personal Comptime Account for: " & user & ControlChars.NewLine &
                                            ControlChars.NewLine &
                                            heading & ControlChars.NewLine &
                                            "------------" & Strings.Space(10) &
                                            "--------------" & Strings.Space(5) &
                                            "---------" & Strings.Space(10) &
                                            "--------" & Strings.Space(9) &
                                            "-------" & ControlChars.NewLine, True)

        My.Computer.FileSystem.WriteAllText(cpath, curdate & Strings.Space(12) &
                                            caseno.PadRight(15, " ") & Strings.Space(6) &
                                            "0.00".PadLeft(5, " ") &
                                            Strings.Space(13) &
                                            "0.00".PadLeft(5, " ") & Strings.Space(13) &
                                            Convert.ToString(previous).PadLeft(5, " ") &
                                            ControlChars.NewLine, True)

        Separation()


    End Sub

    Public Sub PreviewCalculations()

        'declare calculation variables
        Dim earned As Decimal
        Dim calcearned As Decimal
        Dim taken As Decimal
        Dim isEarned As Boolean
        Dim isTaken As Boolean
        Dim previewbankbal As Decimal

        'Determine if this time is accrued or taken
        If accruedRadioButton.Checked Then
            If takenTextBox.Text = String.Empty Then
                takenTextBox.Text = "0.00"
            End If

            'Convert Input
            isEarned = Decimal.TryParse(earnedTextBox.Text, earned)
            isTaken = Decimal.TryParse(takenTextBox.Text, taken)

            'If conversions successful, make calculations
            If isEarned And isTaken Then
                calcearned = earned * 1.5D
                calcearned = Math.Round(calcearned, 2)
                previewbankbal = calcearned + Convert.ToDecimal(prevbalLabel.Text)
                calcearnedTextBox.Text = ""
                calcearnedTextBox.Text = "Total accrued time to enter on affidavit = " &
                                         (earned * 1.5D).ToString("N2") &
                                         " hours" & ControlChars.NewLine &
                                         "-".PadLeft(83, "-") & ControlChars.NewLine &
                                         "Preview of Entry to Activity Sheet:" & ControlChars.NewLine &
                                         ControlChars.NewLine &
                                         "Date Entered" & Strings.Space(13) &
                                         "CaseNo." & Strings.Space(16) &
                                         "Earned(+)" & Strings.Space(12) &
                                         "Taken(-)" & Strings.Space(15) &
                                         "Balance" & ControlChars.NewLine &
                                         "-----------------" & Strings.Space(13) &
                                         "----------" & Strings.Space(16) &
                                         "------------" & Strings.Space(13) &
                                         "----------" & Strings.Space(17) &
                                         "----------" & ControlChars.NewLine &
                                         accruedDateTimePicker.Text & Strings.Space(15) &
                                         caseComboBox.Text.PadRight(15, " ") & Strings.Space(9) &
                                         earnedTextBox.Text.PadLeft(5, " ") & Strings.Space(21) &
                                         takenTextBox.Text.PadLeft(5, " ") & Strings.Space(22) &
                                         Convert.ToString(previewbankbal).PadLeft(5, " ")

                newbalance = calcearned - taken
                newbalance = Math.Round(newbalance, 2)
                newbalLabel.Text = Convert.ToString(newbalance)

            Else : MessageBox.Show("Must be numeric", title, MessageBoxButtons.OK,
                                    MessageBoxIcon.Information)
                earnedTextBox.Focus()
            End If

        ElseIf spentRadioButton.Checked Then

            If earnedTextBox.Text = String.Empty Then
                earnedTextBox.Text = "0.00"
            End If

            'Convert input
            isEarned = Decimal.TryParse(earnedTextBox.Text, earned)
            isTaken = Decimal.TryParse(takenTextBox.Text, taken)

            'If conversions successful, make calculations
            If isEarned And isTaken Then
                calcearned = earned * 1.5D
                calcearned = Math.Round(calcearned, 2)
                newbalance = calcearned - taken
                newbalance = Math.Round(newbalance, 2)
                previewbankbal = newbalance + Convert.ToDecimal(prevbalLabel.Text)
                newbalLabel.Text = Convert.ToString(newbalance)
                calcearnedTextBox.Text = ""
                calcearnedTextBox.Text = "Total taken time to enter on affidavit = " &
                                         (taken).ToString("N2") &
                                         " hours" & ControlChars.NewLine & "-".PadLeft(83, "-") &
                                         ControlChars.NewLine &
                                         "Preview of Entry to Activity Sheet:" & ControlChars.NewLine &
                                         ControlChars.NewLine &
                                         "Date Entered" & Strings.Space(13) &
                                         "CaseNo." & Strings.Space(16) &
                                         "Earned(+)" & Strings.Space(12) &
                                         "Taken(-)" & Strings.Space(15) &
                                         "Balance" &
                                         ControlChars.NewLine &
                                         "-----------------" & Strings.Space(13) &
                                         "----------" & Strings.Space(16) &
                                         "------------" & Strings.Space(13) &
                                         "----------" & Strings.Space(17) &
                                         "----------" & ControlChars.NewLine &
                                         accruedDateTimePicker.Text & Strings.Space(15) &
                                         caseComboBox.Text.PadRight(15, " ") & Strings.Space(9) &
                                         earnedTextBox.Text.PadLeft(5, " ") & Strings.Space(21) &
                                         takenTextBox.Text.PadLeft(5, " ") & Strings.Space(22) &
                                         Convert.ToString(previewbankbal)

            Else : MessageBox.Show("Must be numeric", title, MessageBoxButtons.OK,
            MessageBoxIcon.Information)
                takenTextBox.Focus()
            End If
        End If

        Me.applyButton.Enabled = True
        Me.ApplyToolStripMenuItem.Enabled = True

    End Sub

    Public Sub CalcClear()

        'clears the form

        'declares variables
        Dim my_choice As DialogResult

        my_choice = MessageBox.Show("Do you wish to add to new balance to the bank?", title, MessageBoxButtons.YesNo,
        MessageBoxIcon.Question)
        If my_choice = Windows.Forms.DialogResult.Yes Then
            'declare block variables
            Dim curdate As String
            Dim caseno As String
            Dim line2 As String

            'make calculations
            newbalance = Convert.ToDecimal(newbalLabel.Text)
            newbalance = Math.Round(newbalance, 2)
            previous = Convert.ToDecimal(prevbalLabel.Text)
            previous += newbalance
            previous = Math.Round(previous, 2)
            prevbalLabel.Text = Convert.ToString(previous)

            'convert data
            caseno = caseComboBox.Text
            curdate = accruedDateTimePicker.Text
            line2 = Convert.ToString(previous)

            'Write current balance text file
            If my_choice = Windows.Forms.DialogResult.Yes And My.Computer.FileSystem.FileExists(cpath) Then
                My.Computer.FileSystem.WriteAllText(cpath, curdate & Strings.Space(12) &
                                                    caseno.PadRight(15, " ") & Strings.Space(6) &
                                                    earnedTextBox.Text.PadLeft(5, " ") & Strings.Space(13) &
                                                    takenTextBox.Text.PadLeft(5, " ") & Strings.Space(13) &
                                                    Convert.ToString(previous) & ControlChars.NewLine, True)

                Separation()

            Else

                'Setting up for the first time
                My.Computer.FileSystem.CreateDirectory(cdirectory)
                My.Computer.FileSystem.WriteAllText(cpath, heading & ControlChars.NewLine &
                                                    "------------" & Strings.Space(10) &
                                                    "-------" & Strings.Space(5) &
                                                    "---------" & Strings.Space(10) &
                                                    "--------" & Strings.Space(10) &
                                                    "-------" & ControlChars.NewLine, True)

                My.Computer.FileSystem.WriteAllText(cpath, curdate & Strings.Space(12) &
                                                    caseno.PadRight(15, " ") & Strings.Space(6) &
                                                    earnedTextBox.Text.PadLeft(5, " ") & Strings.Space(13) &
                                                    takenTextBox.Text.PadLeft(5, " ") & Strings.Space(13) &
                                                    Convert.ToString(previous) & ControlChars.NewLine, True)

                Separation()

            End If

            CleanHouse()

        Else : my_choice = Windows.Forms.DialogResult.No
            Me.Show()

            CleanHouse()


        End If
    End Sub

    Private Sub exitApp()
        'Exits the Program

        Dim my_result As DialogResult

        my_result = MessageBox.Show("Are you sure that you are ready to exit?", title,
        MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If my_result = Windows.Forms.DialogResult.No Then

            Me.Show()

            CleanHouse()

        Else : my_result = Windows.Forms.DialogResult.Yes

            If My.Computer.FileSystem.FileExists(cpath) Then
                Me.Close()
            Else : CreateMyPaths()
                Me.Close()
            End If
        End If
    End Sub

    Private Sub Separation()

        My.Computer.FileSystem.WriteAllText(cpath, "".PadLeft(85, "-") &
                                            ControlChars.NewLine, True)

    End Sub

    Private Sub CleanHouse()

        newbalance = 0D
        newbalLabel.Text = "0.00"
        calcearnedTextBox.Text = ""
        calcearnedTextBox.Text = "Ready"
        caseComboBox.Text = ""
        caseComboBox.SelectedItem = "[Enter One]"
        earnedTextBox.Clear()
        takenTextBox.Clear()
        accruedRadioButton.Select()
        accruedDateTimePicker.Focus()
        Me.applyButton.Enabled = False

    End Sub

    '---------------------------------------Buttons and Click Events -------------------------------------------------------------

    Private Sub clearButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles clearButton.Click

        Call CalcClear()

    End Sub

    Private Sub calcButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles calcButton.Click

        Call PreviewCalculations()

    End Sub

    Private Sub exitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles exitButton.Click

        exitApp()

    End Sub

    Private Sub btn_ReconcileData_Click(sender As System.Object, e As System.EventArgs) Handles btn_ReconcileData.Click

        Me.Hide()
        frm_Reconcile.Show()

    End Sub

    Private Sub ClearToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearToolStripMenuItem.Click
        Call CalcClear()
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

        Call ApplyCalculations()

    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click

        exitApp()

    End Sub

    Private Sub ComptimerunToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComptimerunToolStripMenuItem.Click
        'Opens your comptime activity sheet file

        Dim proc As New System.Diagnostics.Process
        proc.StartInfo.FileName = "notepad.exe"
        proc.StartInfo.Arguments = "C:/Comptime/Comptimerun.txt"

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

        Call CalcClear()

    End Sub

    Private Sub CalculateToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CalculateToolStripMenuItem.Click

        Call PreviewCalculations()

    End Sub

    Private Sub ApplyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ApplyToolStripMenuItem.Click

        Call ApplyCalculations()

    End Sub

End Class