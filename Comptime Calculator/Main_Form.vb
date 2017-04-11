'Title                  Comptime Calculator
'Purpose                To calculate comptime time earned or spent
'                       in a particular instance
'Created By             Shon Garrison, December 2008
'Updated Last           August 2016

'Update Notes:          Cosmetic Changes, Updated About Form, Added email provider options, 
'                       Repaired listbox lineup on reconcile sheet


Option Explicit On
Imports System.Globalization


Public Class frm_Main

    Private title As String = "Comptime Calculator"
    Public user As String
    Private newbalance As Decimal
    Private previous As Decimal
    Private text1 As String


    Private Sub clearButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles clearButton.Click

        Call CalcClear()

    End Sub

    Private Sub calcButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles calcButton.Click

        Call PreviewCalculations()

    End Sub


    Private Sub exitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles exitButton.Click

        exitApp()


    End Sub



    Private Sub compcalcForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


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


        'verifies that .txt file exists.  If is does, it assigns it to a txt variable
        'otherwise it exits the program.
        Dim path As String = "C:\Comptime\"
        Dim button As DialogResult

        'sets up a default selection for the radio buttons
        accruedRadioButton.Select()

        'Verify txt file exists else ask if the user wants to create it.
        If My.Computer.FileSystem.FileExists(path & "comptime.txt") Then
            prevbalLabel.Text = My.Computer.FileSystem.ReadAllText(path & "comptime.txt")

            'pulls user variable from text file
            'declare block variables
            Dim readtxt As String
            Dim entry As String
            Dim newLineIndex As Integer = 0
            Dim entryIndex As Integer = 0
            Dim entryuser As String

            'checks for existing comptimerun.txt
            'if it exists, it pulls it and stores it
            If My.Computer.FileSystem.FileExists(path & "comptimerun.txt") Then

                readtxt = My.Computer.FileSystem.ReadAllText(path & "comptimerun.txt")

                'primer for first read of readtxt
                newLineIndex = readtxt.IndexOf(ControlChars.NewLine, entryIndex)

                Do Until newLineIndex = -1

                    'get each line
                    entry = readtxt.Substring(entryIndex, newLineIndex - entryIndex)

                    'finds line  with username that you are searching for
                    entryuser = entry.Contains("Account")

                    'if found, it adds it to preview
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
            button = MessageBox.Show _
            ("The current comptime balance file does not exist.  This is your comptime bank, would you like to create it?", _
            "Comptime Calculator", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)
            If button = DialogResult.Yes Then
                Dim button3 As DialogResult
                button3 = MessageBox.Show("Do you have an current balance to enter?", title, MessageBoxButtons.YesNo, _
                MessageBoxIcon.Question)
                If button3 = Windows.Forms.DialogResult.Yes Then
                    Do Until IsNumeric(prevbalLabel.Text) Or prevbalLabel.Text <> String.Empty
                        prevbalLabel.Text = InputBox("Please enter current balance or click 'Ok' to go to calculator.", title, "0.00")
                        If Not IsNumeric(prevbalLabel.Text) Then
                            MessageBox.Show("Number must be numeric.", title, MessageBoxButtons.OK)
                        End If
                    Loop

                    user = InputBox("Please Enter Your name", title, )

                    Me.Show()
                    Me.Text = "Personal Comptime Calculator for " & user
                    newbalLabel.Text = "0.00"
                    calcearnedTextBox.Text = "Ready"
                ElseIf button3 = Windows.Forms.DialogResult.No Then
                    Me.Show()
                    prevbalLabel.Text = "0.00"
                    newbalLabel.Text = "0.00"
                End If
            Else : button = DialogResult.No
                Me.Close()
            End If
        End If



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

    Private Sub BankToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BankToolStripMenuItem.Click
        'Opends your comptime bank file

        Dim proc As New System.Diagnostics.Process
        proc.StartInfo.FileName = "notepad.exe"
        proc.StartInfo.Arguments = "C:/Comptime/Comptime.txt"

        proc.Start()
    End Sub


    Private Sub AboutToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        frm_About.ShowDialog()
    End Sub

    Private Sub ReadMeFileToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ReadMeFileToolStripMenuItem.Click

        'Opends your comptime bank file

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

    Private Sub btn_Email_Click(sender As System.Object, e As System.EventArgs) Handles btn_Email.Click
        Me.Hide()
        frm_Email.Show()


    End Sub

    Public Sub ApplyCalculations()
        'Saves current balance to txt file

        Dim button3 As DialogResult
        Dim button4 As DialogResult


        'It will ask to you append file.
        button3 = MessageBox.Show("Do you wish to add to new balance to bank?", title, MessageBoxButtons.YesNo, _
        MessageBoxIcon.Question)
        If button3 = Windows.Forms.DialogResult.Yes Then

            'make calculations
            newbalance = Convert.ToDecimal(newbalLabel.Text)
            newbalance = Math.Round(newbalance, 2)
            previous = Convert.ToDecimal(prevbalLabel.Text)
            previous = previous + newbalance
            previous = Math.Round(previous, 2)
            prevbalLabel.Text = Convert.ToString(previous)

            'Declare text writing variables
            Dim path1 As String = "C:\Comptime\comptime.txt"
            Dim path2 As String = "C:\Comptime\"
            Dim line As String
            Dim caseno As String
            Dim curdate As String
            Dim heading As String = "Date Entered" & Strings.Space(10) & "CaseNo." & Strings.Space(13) & _
            "Earned(+)" & Strings.Space(10) & "Taken(-)" & Strings.Space(9) & "Balance"

            'Convert the data from the Previous Balance Label and store it in line variable
            line = Convert.ToString(previous)
            caseno = caseComboBox.Text
            curdate = accruedDateTimePicker.Text

            'If Comptime file exists, the prog writes current balance text file
            If My.Computer.FileSystem.FileExists(path2 & "comptimerun.txt") And _
            My.Computer.FileSystem.FileExists(path1) Then
                My.Computer.FileSystem.WriteAllText(path1, line, False)
                My.Computer.FileSystem.WriteAllText(path2 & "comptimerun.txt", _
                curdate & Strings.Space(12) & _
                caseno.PadRight(15, " ") & Strings.Space(6) & _
                earnedTextBox.Text.PadLeft(5, " ") & Strings.Space(13) & _
                takenTextBox.Text.PadLeft(5, " ") & Strings.Space(13) & _
                Convert.ToString(previous).PadLeft(5, " ") & ControlChars.NewLine, True)

                My.Computer.FileSystem.WriteAllText(path2 & "comptimerun.txt", "".PadLeft(85, "-") _
                & ControlChars.NewLine, True)

                MessageBox.Show("Processing complete. The form will be cleared.", _
                                title, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.Show()
                newbalLabel.Text = "0.00"
                calcearnedTextBox.Text = "Ready"
                caseComboBox.Text = ""
                earnedTextBox.Clear()
                takenTextBox.Clear()
                accruedRadioButton.Select()
                newbalance = 0D
                accruedDateTimePicker.Focus()
                Me.applyButton.Enabled = False

                'If Comptime file does not exists, the prog creates it and 
                'writes current balance text file
            Else : My.Computer.FileSystem.CreateDirectory("C:\Comptime")
                My.Computer.FileSystem.WriteAllText(path1, line, False)
                My.Computer.FileSystem.WriteAllText(path2 & "comptimerun.txt", _
                    "Orange County Juvenile Probation Dept" & ControlChars.NewLine & _
                    "---------------------------------------" & ControlChars.NewLine & _
                    "Personal Comptime Account for: " & user & ControlChars.NewLine & _
                    heading & ControlChars.NewLine & _
                    "------------" & Strings.Space(10) & _
                    "--------------" & Strings.Space(5) & _
                    "---------" & Strings.Space(10) & _
                    "--------" & Strings.Space(9) & _
                    "-------" & ControlChars.NewLine, True)

                My.Computer.FileSystem.WriteAllText(path2 & "comptimerun.txt", _
                    curdate & Strings.Space(12) & _
                    caseno.PadRight(15, " ") & Strings.Space(6) & _
                    earnedTextBox.Text.PadLeft(5, " ") & Strings.Space(13) & _
                    takenTextBox.Text.PadLeft(5, " ") & Strings.Space(13) & _
                    Convert.ToString(previous).PadLeft(5, " ") & ControlChars.NewLine, True)

                My.Computer.FileSystem.WriteAllText(path2 & "comptimerun.txt", "".PadLeft(85, "-") & _
                                                    ControlChars.NewLine, True)

                MessageBox.Show("Processing complete. The form will be cleared.", _
                                title, MessageBoxButtons.OK, MessageBoxIcon.Information)

                Me.Show()
                newbalLabel.Text = "0.00"
                calcearnedTextBox.Text = "Ready"
                caseComboBox.Text = ""
                earnedTextBox.Clear()
                takenTextBox.Clear()
                accruedRadioButton.Select()
                newbalance = 0D
                accruedDateTimePicker.Focus()
                Me.applyButton.Enabled = False

            End If

            'If user does not want to make calculation, the program
            'will ask if the user wants to return to the program for another calculation.
            'If the user does not, then the program will direct user to exit.
        Else : button3 = Windows.Forms.DialogResult.No
            button4 = MessageBox.Show("Do you want to make another calculation?", title, _
        MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If button4 = Windows.Forms.DialogResult.Yes Then
                Me.Show()
                newbalLabel.Text = "0.00"
                calcearnedTextBox.Text = "Ready"
                caseComboBox.Text = ""
                earnedTextBox.Clear()
                takenTextBox.Clear()
                accruedRadioButton.Select()
                newbalance = 0D
                accruedDateTimePicker.Focus()
                Me.applyButton.Enabled = False

            Else : button4 = Windows.Forms.DialogResult.No
                MessageBox.Show("No calcuation will be made and the form will be reset. You may exit the program.", title, _
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.Show()
                newbalLabel.Text = "0.00"
                calcearnedTextBox.Text = "Ready"
                caseComboBox.Text = ""
                earnedTextBox.Clear()
                takenTextBox.Clear()
                accruedRadioButton.Select()
                newbalance = 0D
                accruedDateTimePicker.Focus()
                Me.applyButton.Enabled = False
            End If

        End If

    End Sub

    Public Sub PreviewCalculations()

        'declare variables
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
                calcearnedTextBox.Text = "Total accrued time to enter on affidavit = " & _
                (earned * 1.5D).ToString("N2") & " hours" & ControlChars.NewLine & _
                "-".PadLeft(83, "-") & ControlChars.NewLine & _
                "Preview of Entry to Activity Sheet:" & ControlChars.NewLine & ControlChars.NewLine & _
                "Date Entered" & Strings.Space(13) & _
                "CaseNo." & Strings.Space(16) & _
                "Earned(+)" & Strings.Space(12) & _
                "Taken(-)" & Strings.Space(15) & _
                "Balance" & ControlChars.NewLine & _
                "-----------------" & Strings.Space(13) & _
                "----------" & Strings.Space(16) & _
                "------------" & Strings.Space(13) & _
                "----------" & Strings.Space(17) & _
                "----------" & ControlChars.NewLine & _
                accruedDateTimePicker.Text & Strings.Space(15) & _
                caseComboBox.Text.PadRight(15, " ") & Strings.Space(9) & _
                earnedTextBox.Text.PadLeft(5, " ") & Strings.Space(21) & _
                takenTextBox.Text.PadLeft(5, " ") & Strings.Space(22) & _
                Convert.ToString(previewbankbal).PadLeft(5, " ")

                newbalance = calcearned - taken
                newbalance = Math.Round(newbalance, 2)
                newbalLabel.Text = Convert.ToString(newbalance)

            Else : MessageBox.Show("Must be numeric", title, MessageBoxButtons.OK, _
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
                calcearnedTextBox.Text = "Total taken time to enter on affidavit = " & _
                (taken).ToString("N2") & _
                " hours" & ControlChars.NewLine & _
                "-".PadLeft(83, "-") & ControlChars.NewLine & _
                "Preview of Entry to Activity Sheet:" & ControlChars.NewLine & ControlChars.NewLine & _
                "Date Entered" & Strings.Space(13) & _
                "CaseNo." & Strings.Space(16) & _
                "Earned(+)" & Strings.Space(12) & _
                "Taken(-)" & Strings.Space(15) & _
                "Balance" & ControlChars.NewLine & _
                "-----------------" & Strings.Space(13) & _
                "----------" & Strings.Space(16) & _
                "------------" & Strings.Space(13) & _
                "----------" & Strings.Space(17) & _
                "----------" & ControlChars.NewLine & _
                accruedDateTimePicker.Text & Strings.Space(15) & _
                caseComboBox.Text.PadRight(15, " ") & Strings.Space(9) & _
                earnedTextBox.Text.PadLeft(5, " ") & Strings.Space(21) & _
                takenTextBox.Text.PadLeft(5, " ") & Strings.Space(22) & _
                Convert.ToString(previewbankbal)

            Else : MessageBox.Show("Must be numeric", title, MessageBoxButtons.OK, _
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
        Dim button As DialogResult

        button = MessageBox.Show("Do you wish to add to new balance to the bank?", title, MessageBoxButtons.YesNo, _
        MessageBoxIcon.Question)
        If button = Windows.Forms.DialogResult.Yes Then
            'declare block variables
            Dim curdate As String
            Dim caseno As String
            Dim line2 As String
            Dim heading As String = "Date Entered" & Strings.Space(10) & "CaseNo." & Strings.Space(5) & _
            "Earned(+)" & Strings.Space(10) & "Taken(-)" & Strings.Space(10) & "Balance"
            Dim path3 As String = "C:\Comptime\"
            Dim path4 As String = "C:\Comptime\comptime.txt"

            'make calculations
            newbalance = Convert.ToDecimal(newbalLabel.Text)
            newbalance = Math.Round(newbalance, 2)
            previous = Convert.ToDecimal(prevbalLabel.Text)
            previous = previous + newbalance
            previous = Math.Round(previous, 2)
            prevbalLabel.Text = Convert.ToString(previous)

            'convert data
            caseno = caseComboBox.Text
            curdate = accruedDateTimePicker.Text
            line2 = Convert.ToString(previous)

            'Write current balance text file
            If button = Windows.Forms.DialogResult.Yes And My.Computer.FileSystem.FileExists(path3 & "comptimerun.txt") _
            And My.Computer.FileSystem.FileExists(path4) Then
                My.Computer.FileSystem.WriteAllText(path4, line2, False)
                My.Computer.FileSystem.WriteAllText(path3 & "comptimerun.txt", _
                    curdate & Strings.Space(12) & _
                    caseno.PadRight(15, " ") & Strings.Space(6) & _
                    earnedTextBox.Text.PadLeft(5, " ") & Strings.Space(13) & _
                    takenTextBox.Text.PadLeft(5, " ") & Strings.Space(13) & _
                    Convert.ToString(previous) & ControlChars.NewLine, True)

                My.Computer.FileSystem.WriteAllText(path3 & "comptimerun.txt", "".PadLeft(85, "-") & ControlChars.NewLine, _
                True)

            Else 'Setting up for the first time
                My.Computer.FileSystem.CreateDirectory("C:\Comptime")
                My.Computer.FileSystem.WriteAllText(path4, line2, False)
                My.Computer.FileSystem.WriteAllText(path3 & "comptimerun.txt", _
                    heading & ControlChars.NewLine & _
                    "------------" & Strings.Space(10) & _
                    "-------" & Strings.Space(5) & _
                    "---------" & Strings.Space(10) & _
                    "--------" & Strings.Space(10) & _
                    "-------" & ControlChars.NewLine, True)

                My.Computer.FileSystem.WriteAllText(path3 & "comptimerun.txt", _
                    curdate & Strings.Space(12) & _
                    caseno.PadRight(15, " ") & Strings.Space(6) & _
                    earnedTextBox.Text.PadLeft(5, " ") & Strings.Space(13) & _
                    takenTextBox.Text.PadLeft(5, " ") & Strings.Space(13) & _
                    Convert.ToString(previous) & ControlChars.NewLine, True)

                My.Computer.FileSystem.WriteAllText(path3 & "comptimerun.txt", "".PadLeft(85, "-") & ControlChars.NewLine, _
                True)
            End If

            calcearnedTextBox.Text = ""
            calcearnedTextBox.Text = "Ready"
            newbalLabel.Text = "0.00"
            earnedTextBox.Clear()
            caseComboBox.SelectedItem = "[Enter One]"
            takenTextBox.Clear()
            accruedRadioButton.Select()
            accruedDateTimePicker.Focus()
            Me.applyButton.Enabled = False

        Else : button = Windows.Forms.DialogResult.No
            Me.Show()
            newbalance = 0D
            newbalLabel.Text = "0.00"
            calcearnedTextBox.Text = ""
            calcearnedTextBox.Text = "Ready"
            caseComboBox.SelectedItem = "[Enter One]"
            earnedTextBox.Clear()
            takenTextBox.Clear()
            accruedRadioButton.Select()
            accruedDateTimePicker.Focus()
            Me.applyButton.Enabled = False

        End If
    End Sub
    
    Private Sub ClearToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearToolStripMenuItem.Click
        Call CalcClear()
    End Sub


    Private Sub exitApp()
        'Exits the Program

        Dim button2 As DialogResult

        button2 = MessageBox.Show("Are you sure that you are ready to exit?", title, _
        MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If button2 = Windows.Forms.DialogResult.No Then
            Me.Show()
            newbalLabel.Text = "0.00"
            calcearnedTextBox.Text = "Ready"
            caseComboBox.SelectedItem = "[Enter One]"
            earnedTextBox.Clear()
            takenTextBox.Clear()
            accruedRadioButton.Select()
            newbalance = 0D
            accruedDateTimePicker.Focus()
            Me.applyButton.Enabled = False

        Else : button2 = Windows.Forms.DialogResult.Yes
            Me.Close()

        End If
    End Sub

    Private Sub btn_ReconcileData_Click(sender As System.Object, e As System.EventArgs) Handles btn_ReconcileData.Click

        Me.Hide()
        frm_Reconcile.Show()

    End Sub
End Class
