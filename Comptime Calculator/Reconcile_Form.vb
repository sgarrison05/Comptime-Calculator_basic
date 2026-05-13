
Public Class frm_Reconcile

    Dim reconcilledPath As String = "D:\Temp\Comptime\" & CStr(Year(Today) - 1)

    Private Sub frm_Reconcile_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        'default selects the previous year
        Me.txtYear.Text = CStr(Year(Today) - 1)
        Me.btnReconcile.Enabled = False
        Me.libxReconcile.Items.Clear()

        Try
            If My.Computer.FileSystem.FileExists(frm_Main.CPATH) Then
                Dim allEntries As List(Of String) = ReadAllTransactionLines(frm_Main.CPATH)
                For Each line As String In allEntries
                    Me.libxReconcile.Items.Add(line)
                Next
            End If
        Catch ex As Exception
            MessageBox.Show("Error loading bank file:" & Environment.NewLine & ex.Message,
                            frm_Main.TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Me.txtYear.Focus()

    End Sub

    Private Sub btnExit_Click(sender As System.Object, e As System.EventArgs) Handles btnExit.Click

        Me.Close()
        frm_Main.Show()


    End Sub

    Private Sub btnPreview_Click(sender As Object, e As EventArgs) Handles btnPreview.Click

        'Previews the Previous Year

        Me.libxPreview.Items.Clear()
        Me.libxOrig.Items.Clear()

        Dim previousYear As Integer
        If Not Integer.TryParse(Me.txtYear.Text, previousYear) Then
            MessageBox.Show("Please enter a valid 4-digit year.", frm_Main.TITLE,
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Me.txtYear.Focus()
            Return
        End If

        Dim currentYear As Integer = previousYear + 1

        Try
            If Not My.Computer.FileSystem.FileExists(frm_Main.CPATH) Then
                MessageBox.Show("Bank file not found. Nothing to preview.",
                                frm_Main.TITLE, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim previousYearEntries As List(Of String) =
                GetEntriesForYear(frm_Main.CPATH, previousYear)

            Dim currentYearEntries As List(Of String) =
                GetEntriesForYear(frm_Main.CPATH, currentYear)

            For Each line As String In previousYearEntries
                Me.libxPreview.Items.Add(line)
            Next

            For Each line As String In currentYearEntries
                Me.libxOrig.Items.Add(line)
            Next

        Catch ex As Exception
            MessageBox.Show("Error reading bank file:" & Environment.NewLine & ex.Message,
                            frm_Main.TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try

        're-enables reconcile button after preview selected
        Me.btnReconcile.Enabled = True

    End Sub

    Private Sub btnReconcile_Click(sender As Object, e As EventArgs) Handles btnReconcile.Click

        'Creates reconcilled text document for previous year 
        'in it's own folder

        If Me.libxPreview.Items.Count = 0 OrElse Me.libxOrig.Items.Count = 0 Then
            MessageBox.Show("Make sure you have a previous year to reconcile.",
                            "Important", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Me.btnClearPrev.Focus()
            Return
        End If

        Dim previousYear As Integer
        Integer.TryParse(Me.txtYear.Text, previousYear)

        Try
            '--- Write reconciled archive file for previous year ---
            Dim archiveDir As String = ArchivePath(previousYear)
            Dim archiveFile As String = ReconciledFilePath(previousYear)

            If Not My.Computer.FileSystem.DirectoryExists(archiveDir) Then
                My.Computer.FileSystem.CreateDirectory(archiveDir)
            End If

            ' Always overwrite the archive file cleanly
            My.Computer.FileSystem.WriteAllText(archiveFile, String.Empty, False)
            WriteReconciledHeader(archiveFile, previousYear)

            For Each item As String In Me.libxPreview.Items
                My.Computer.FileSystem.WriteAllText(archiveFile, item & ControlChars.NewLine, True)
                My.Computer.FileSystem.WriteAllText(archiveFile,
                    "".PadLeft(100, "-"c) & ControlChars.NewLine, True)
            Next

            '--- Rebuild active bank file with current-year entries only ---
            My.Computer.FileSystem.WriteAllText(frm_Main.CPATH, String.Empty, False)
            WriteMainFileHeader(frm_Main.CPATH)

            For Each item As String In Me.libxOrig.Items
                My.Computer.FileSystem.WriteAllText(frm_Main.CPATH,
                    item & ControlChars.NewLine, True)
                My.Computer.FileSystem.WriteAllText(frm_Main.CPATH,
                    "".PadLeft(100, "-"c) & ControlChars.NewLine, True)
            Next

            MessageBox.Show("Reconciliation complete." & Environment.NewLine &
                            "The form will now close and return to the main form.",
                            frm_Main.TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information)

            Me.Close()
            frm_Main.Show()

        Catch ex As Exception
            MessageBox.Show("Error during reconciliation:" & Environment.NewLine & ex.Message,
                            frm_Main.TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub WriteReconciledHeader(filepath As String, year As Integer)

        My.Computer.FileSystem.WriteAllText(filepath,
           "Orange County Juvenile Probation Dept." & ControlChars.NewLine &
           "---------------------------------------" & ControlChars.NewLine &
           "Personal Comptime Account for: " & frm_Main.userName & ControlChars.NewLine &
           ControlChars.NewLine &
           frm_Main._heading & ControlChars.NewLine &
           frm_Main._columnDivider & ControlChars.NewLine, True)

    End Sub

    Private Sub WriteMainFileHeader(filePath As String)
        My.Computer.FileSystem.WriteAllText(filePath,
            "Orange County Juvenile Probation Dept." & ControlChars.NewLine &
            "---------------------------------------" & ControlChars.NewLine &
            "Personal Comptime Account for: " & frm_Main.userName & ControlChars.NewLine &
            ControlChars.NewLine &
            frm_Main._heading & ControlChars.NewLine &
            frm_Main._columnDivider & ControlChars.NewLine, True)
    End Sub

    Private Function ReadAllTransactionLines(filepath As String) As List(Of String)

        Dim results As New List(Of String)
        Dim fileText As String = My.Computer.FileSystem.ReadAllText(filepath)
        Dim entryIndex As Integer = 0
        Dim newLineIndex As Integer = fileText.IndexOf(ControlChars.NewLine, entryIndex)

        Do Until newLineIndex = -1
            Dim line As String = fileText.Substring(entryIndex, newLineIndex - entryIndex)
            If line.Contains("/") Then
                results.Add(line)
            End If
            entryIndex = newLineIndex + 2
            newLineIndex = fileText.IndexOf(ControlChars.NewLine, entryIndex)
        Loop

        Return results

    End Function

    Private Function GetEntriesForYear(Filepath As String, year As Integer) As List(Of String)

        Dim results As New List(Of String)
        Dim yearString As String = year.ToString()
        Dim fileText As String = My.Computer.FileSystem.ReadAllText(Filepath)
        Dim entryIndex As Integer = 0
        Dim newLineIndex As Integer = fileText.IndexOf(ControlChars.NewLine, entryIndex)

        Do Until newLineIndex = -1
            Dim line As String = fileText.Substring(entryIndex, newLineIndex - entryIndex)
            If line.Contains(yearString) Then
                results.Add(line)
            End If
            entryIndex = newLineIndex + 2
            newLineIndex = fileText.IndexOf(ControlChars.NewLine, entryIndex)
        Loop

        Return results

    End Function

    Private Function ArchivePath(year As Integer) As String
        Return System.IO.Path.Combine(frm_Main.CDIRECTORY, year.ToString())
    End Function

    Private Function ReconciledFilePath(year As Integer) As String
        Return System.IO.Path.Combine(ArchivePath(year),
                                      "Comptimerun_Reconciled_" & year.ToString() & ".txt")
    End Function

    Private Sub txtYear_Enter(sender As Object, e As EventArgs) Handles txtYear.Enter

        Me.txtYear.SelectAll()

    End Sub

    Private Sub txtYear_TextChanged(sender As Object, e As EventArgs) Handles txtYear.TextChanged

        Me.libxPreview.Items.Clear()
        Me.libxOrig.Items.Clear()
        Me.btnReconcile.Enabled = False

    End Sub

    Private Sub btnClearPrev_Click(sender As Object, e As EventArgs) Handles btnClearPrev.Click

        Me.libxPreview.Items.Clear()
        Me.libxOrig.Items.Clear()

    End Sub

   
End Class