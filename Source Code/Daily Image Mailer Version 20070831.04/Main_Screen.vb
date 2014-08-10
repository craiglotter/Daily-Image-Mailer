Imports System.IO
Imports System.Net.Mail

Public Class Main_Screen

    Dim progresslabel As String = ""
    Dim attachmentpath As String = ""
    Dim filename As String = ""

    Dim attachmentsendfolder As String = ""
    Dim attachmentsentfolder As String = ""
    Dim attachmentdeletefolder As String = ""

    Dim shownminimizetip As Boolean = False

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message()
                If FullErrors_Checkbox.Checked = True Then
                    Display_Message1.Message_Textbox.Text = "The Application encountered the following problem: " & vbCrLf & identifier_msg & ": " & ex.ToString
                Else
                    Display_Message1.Message_Textbox.Text = "The Application encountered the following problem: " & vbCrLf & identifier_msg & ": " & ex.Message.ToString
                End If
                Display_Message1.Timer1.Interval = 1000
                Display_Message1.ShowDialog()
                Dim dir As System.IO.DirectoryInfo = New System.IO.DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                dir = Nothing
                Dim filewriter As System.IO.StreamWriter = New System.IO.StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & identifier_msg & ": " & ex.ToString)
                filewriter.WriteLine("")
                filewriter.Flush()
                filewriter.Close()
                filewriter = Nothing
                Label2.Text = "Error encountered in last action"
            End If
        Catch exc As Exception
            MsgBox("An error occurred in the application's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub RunWorker()
        Try
            Label2.Text = "Preparing to send mail"
            progresslabel = ""
            ProgressBar1.Enabled = True
            ProgressBar1.Value = 0
            DateTimePicker1.Enabled = False
            DateTimePicker2.Enabled = False
            NumericUpDown1.Enabled = False
            CheckBox1.Enabled = False

            mailserver.Enabled = False
            mailserverport.Enabled = False
            emailaddress.Enabled = False
            filename = ""
            attachmentpath = ""

            BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception
            Error_Handler(ex, "Run Worker")
        End Try
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            shownminimizetip = False
            Control.CheckForIllegalCrossThreadCalls = False
            DateTimePicker1.Value = Now
            DateTimePicker2.Value = Now
            attachmentsendfolder = (Application.StartupPath & "\Attachments").Replace("\\", "\")
            attachmentsentfolder = (Application.StartupPath & "\Attachments Already Sent").Replace("\\", "\")
            attachmentdeletefolder = (Application.StartupPath & "\Attachments To Delete").Replace("\\", "\")
            Me.Text = My.Application.Info.ProductName & " (" & Format(My.Application.Info.Version.Major, "0000") & Format(My.Application.Info.Version.Minor, "00") & Format(My.Application.Info.Version.Build, "00") & "." & Format(My.Application.Info.Version.Revision, "00") & ")"
            If My.Computer.FileSystem.DirectoryExists(attachmentsendfolder) = False Then
                My.Computer.FileSystem.CreateDirectory(attachmentsendfolder)
            End If
            If My.Computer.FileSystem.DirectoryExists(attachmentsentfolder) = False Then
                My.Computer.FileSystem.CreateDirectory(attachmentsentfolder)
            End If
            If My.Computer.FileSystem.DirectoryExists(attachmentdeletefolder) = False Then
                My.Computer.FileSystem.CreateDirectory(attachmentdeletefolder)
            End If
            Label11.Text = (attachmentsendfolder)
            loadSettings()
            EmptyToDeleteFolder()
            Label2.Text = "Application loaded"
            Label7.Select()
        Catch ex As Exception
            Error_Handler(ex, "Form Load")
        End Try
    End Sub

    Private Sub EmptyToDeleteFolder()
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo(attachmentdeletefolder)
            For Each fil As FileInfo In dir.GetFiles
                Try
                    fil.Delete()
                    fil = Nothing
                Catch ex As Exception
                    Error_Handler(ex, "Empty To Delete Folder")
                End Try
            Next
            dir = Nothing
        Catch ex As Exception
            Error_Handler(ex, "Empty To Delete Folder")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            If emailaddress.Text.Length > 0 And mailserver.Text.Length > 0 Then
                If My.Computer.Network.IsAvailable = True Then
                    progresslabel = "Creating server object"
                    BackgroundWorker1.ReportProgress(0)
                    Dim obj As SmtpClient
                    If mailserverport.Text.Length > 0 Then
                        obj = New SmtpClient(mailserver.Text, mailserverport.Text)
                    Else
                        obj = New SmtpClient(mailserver.Text)
                    End If
                    progresslabel = "Creating mail object"
                    BackgroundWorker1.ReportProgress(5)

                    Dim msg As MailMessage = New MailMessage

                    msg.Subject = "Your Daily Update"
                    Dim fromaddress As MailAddress = New MailAddress("unattended-mailbox@obe1.com.uct.ac.za", "Daily Image Mailer")
                    msg.From = fromaddress

                    progresslabel = "Adding recipients"
                    BackgroundWorker1.ReportProgress(15)

                    For Each token As String In emailaddress.Text.Split(";")
                        msg.To.Add(token)
                    Next
                    progresslabel = "Adding attachments"
                    BackgroundWorker1.ReportProgress(25)


                    If My.Computer.FileSystem.DirectoryExists(attachmentsendfolder) = False Then
                        My.Computer.FileSystem.CreateDirectory(attachmentsendfolder)
                    End If
                    If My.Computer.FileSystem.DirectoryExists(attachmentsentfolder) = False Then
                        My.Computer.FileSystem.CreateDirectory(attachmentsentfolder)
                    End If
                    If My.Computer.FileSystem.DirectoryExists(attachmentdeletefolder) = False Then
                        My.Computer.FileSystem.CreateDirectory(attachmentdeletefolder)
                    End If

                    Dim dir As DirectoryInfo = New DirectoryInfo(attachmentsendfolder)

                    If My.Computer.FileSystem.FileExists((attachmentsendfolder & "\Thumbs.db").Replace("\\", "\")) = True Then
                        My.Computer.FileSystem.DeleteFile((attachmentsendfolder & "\Thumbs.db").Replace("\\", "\"), FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                    End If

                    Dim alreadyset As Boolean = False
                    For Each fil As FileInfo In dir.GetFiles()
                        If alreadyset = False Then

                            attachmentpath = fil.FullName
                            filename = fil.Name
                            fil = Nothing
                            alreadyset = True
                            Exit For
                        
                        End If
                    Next


                    If attachmentpath.Length < 1 Then

                        Dim dir2 As DirectoryInfo = New DirectoryInfo(attachmentsentfolder)

                        For Each fil2 As FileInfo In dir2.GetFiles()
                            If Not fil2.Name = "Thumbs.db" Then
                                My.Computer.FileSystem.MoveFile(fil2.FullName, (attachmentsendfolder & "\" & fil2.Name).Replace("\\", "\"), True)
                            End If
                            fil2 = Nothing
                        Next
                        dir2 = Nothing


                        alreadyset = False
                        For Each fil As FileInfo In dir.GetFiles()
                            If alreadyset = False Then
                                attachmentpath = fil.FullName
                                filename = fil.Name
                                fil = Nothing
                                alreadyset = True
                                Exit For
                            End If
                        Next

                    End If
                    dir = Nothing




                    If attachmentpath.Length > 0 Then
                        My.Computer.FileSystem.CopyFile(attachmentpath, (attachmentsentfolder & "\" & filename).Replace("\\", "\"), True)
                        Dim adjustedfilename As String = filename
                        Dim counter As Integer = 1
                        While My.Computer.FileSystem.FileExists((attachmentdeletefolder & "\" & adjustedfilename).Replace("\\", "\")) = True
                            adjustedfilename = adjustedfilename & "(" & counter & ")"
                            counter = counter + 1
                        End While
                        My.Computer.FileSystem.MoveFile(attachmentpath, (attachmentdeletefolder & "\" & adjustedfilename).Replace("\\", "\"), True)
                        attachmentpath = (attachmentdeletefolder & "\" & filename).Replace("\\", "\")
                        Dim att As Attachment = New Attachment(attachmentpath)
                        msg.Attachments.Add(att)
                        progresslabel = "Sending mail"
                        BackgroundWorker1.ReportProgress(55)
                        obj.Send(msg)
                        att = Nothing
                        progresslabel = "Mail successfully sent"
                        BackgroundWorker1.ReportProgress(100)
                    Else
                        e.Cancel = True
                        progresslabel = "No attachments to send out"
                        BackgroundWorker1.ReportProgress(100)
                    End If
                    msg = Nothing
                    obj = Nothing

                 



                Else
                    e.Cancel = True
                    progresslabel = "No available network"
                    BackgroundWorker1.ReportProgress(100)
                End If
            Else
                e.Cancel = True
                progresslabel = "No user input"
                BackgroundWorker1.ReportProgress(100)
            End If
        Catch ex As Exception
            Error_Handler(ex, "Send Email")
            progresslabel = "Failed to send mail: Error reported"
        End Try
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As System.Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Try
            ProgressBar1.Value = e.ProgressPercentage
            Label2.Text = progresslabel
        Catch ex As Exception
            Error_Handler(ex, "Worker Progress Changed")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            ProgressBar1.Value = 100
            ProgressBar1.Enabled = False

            CheckBox1.Enabled = True
            If CheckBox1.Checked = True Then
                NumericUpDown1.Enabled = True
                DateTimePicker1.Enabled = False
                DateTimePicker2.Enabled = False
            Else
                NumericUpDown1.Enabled = False
                DateTimePicker1.Enabled = True
                DateTimePicker2.Enabled = True
            End If

            mailserver.Enabled = True
            mailserverport.Enabled = True
            emailaddress.Enabled = True

            If e.Cancelled = True Then
                Label2.Text = "Failed to send the mail: " & progresslabel
            Else
                Label2.Text = progresslabel
            End If

        Catch ex As Exception
            Error_Handler(ex, "Run Worker Completed")
        End Try
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        Try
            Label2.Text = "About displayed"
            AboutBox1.ShowDialog()
        Catch ex As Exception
            Error_Handler(ex, "Display About Screen")
        End Try
    End Sub

    Private Sub HelpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripMenuItem.Click
        Try
            Label2.Text = "Help displayed"
            HelpBox1.ShowDialog()
        Catch ex As Exception
            Error_Handler(ex, "Display Help Screen")
        End Try
    End Sub

   

    Private Sub SaveSettings()
        Try
            My.Settings.Save()
            My.Settings.EmailAddress = emailaddress.Text
            My.Settings.MailServer = mailserver.Text
            My.Settings.MailServerPort = mailserverport.Text
            My.Settings.TimeOfDayHour = Format(DateTimePicker1.Value, "HH")
            My.Settings.TimeOfDayMinute = Format(DateTimePicker1.Value, "mm")
            My.Settings.TimeOfDaySecond = Format(DateTimePicker1.Value, "ss")
            My.Settings.TimeOfDayHour2 = Format(DateTimePicker2.Value, "HH")
            My.Settings.TimeOfDayMinute2 = Format(DateTimePicker2.Value, "mm")
            My.Settings.TimeOfDaySecond2 = Format(DateTimePicker2.Value, "ss")
            My.Settings.TimeInterval = NumericUpDown1.Value
            My.Settings.TimeIntervalSelected = CheckBox1.Checked
            My.Settings.Save()
        Catch ex As Exception
            Error_Handler(ex, "Save Settings")
        End Try
    End Sub

    Private Sub loadSettings()
        Try
            If Not My.Settings Is Nothing Then
                If Not My.Settings.EmailAddress Is Nothing Then
                    emailaddress.Text = My.Settings.EmailAddress
                End If
                If Not My.Settings.MailServer Is Nothing Then
                    mailserver.Text = My.Settings.MailServer
                End If
                If Not My.Settings.MailServerPort Is Nothing Then
                    mailserverport.Text = My.Settings.MailServerPort
                End If
                If Not My.Settings.TimeInterval Is Nothing Then
                    If My.Settings.TimeInterval.Length > 0 Then
                        NumericUpDown1.Value = Integer.Parse(My.Settings.TimeInterval)
                    End If
                End If
                'If Not My.Settings.TimeIntervalSelected Is Nothing Then
                '    If My.Settings.TimeIntervalSelected.Length > 0 Then
                '        If (My.Settings.TimeIntervalSelected = "False") Or (My.Settings.TimeIntervalSelected = "True") Then
                '            CheckBox1.Checked = Boolean.Parse(My.Settings.TimeIntervalSelected)
                '        End If
                '    End If
                'End If

                If Not My.Settings.TimeOfDayHour Is Nothing And Not My.Settings.TimeOfDayMinute Is Nothing And Not My.Settings.TimeOfDaySecond Is Nothing Then
                    Dim dd As Date
                    If Not "01/01/2001 " & My.Settings.TimeOfDayHour & ":" & My.Settings.TimeOfDayMinute & ":" & My.Settings.TimeOfDaySecond = "01/01/2001 " & ":" & ":" Then
                        dd = Date.Parse("01/01/2001 " & My.Settings.TimeOfDayHour & ":" & My.Settings.TimeOfDayMinute & ":" & My.Settings.TimeOfDaySecond)
                        DateTimePicker1.Value = dd
                    End If
                    dd = Nothing
                End If
                If Not My.Settings.TimeOfDayHour2 Is Nothing And Not My.Settings.TimeOfDayMinute2 Is Nothing And Not My.Settings.TimeOfDaySecond2 Is Nothing Then
                    Dim dd As Date
                    If Not "01/01/2001 " & My.Settings.TimeOfDayHour2 & ":" & My.Settings.TimeOfDayMinute2 & ":" & My.Settings.TimeOfDaySecond2 = "01/01/2001 " & ":" & ":" Then
                        dd = Date.Parse("01/01/2001 " & My.Settings.TimeOfDayHour2 & ":" & My.Settings.TimeOfDayMinute2 & ":" & My.Settings.TimeOfDaySecond2)
                        DateTimePicker2.Value = dd
                    End If
                    dd = Nothing
                End If


            End If
            'default values
            If emailaddress.Text.Length < 1 Then
                emailaddress.Text = "Craig.Lotter@uct.ac.za"
            End If
            If mailserver.Text.Length < 1 Then
                mailserver.Text = "obe1.com.uct.ac.za"
            End If
            If mailserverport.Text.Length < 1 Then
                mailserverport.Text = "25"
            End If
        Catch ex As Exception
            Error_Handler(ex, "Load Settings")
        End Try
    End Sub

    Private Sub Main_Screen_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Try
            CheckBox1.Checked = False
            SaveSettings()
        Catch ex As Exception
            Error_Handler(ex, "Closing Application")
        End Try
    End Sub


    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            Label7.Text = Format(Now, "HH:mm:ss")
            If Label7.Text = Label8.Text Or Label7.Text = Label1.Text Then
                If Label7.Text = Label1.Text And CheckBox1.Checked = True Then
                    'ignore because we're running on the scheduled timer
                Else
                    RunWorker()
                    If CheckBox1.Checked = True Then
                        DateTimePicker1.Value = DateTimePicker1.Value.AddMinutes(NumericUpDown1.Value)
                    End If
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex, "Timer Ticking")
        End Try
    End Sub


    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        Try
            Label8.Text = Format(DateTimePicker1.Value, "HH:mm:ss")
        Catch ex As Exception
            Error_Handler(ex, "Change Scheduled Time")
        End Try
    End Sub

    Private Sub DateTimePicker2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker2.ValueChanged
        Try
            Label1.Text = Format(DateTimePicker2.Value, "HH:mm:ss")
        Catch ex As Exception
            Error_Handler(ex, "Change Scheduled Time")
        End Try
    End Sub


    Private Sub NotifyIcon1_BalloonTipClicked(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NotifyIcon1.BalloonTipClicked
        Try
            Me.WindowState = FormWindowState.Normal
            Me.ShowInTaskbar = True
            NotifyIcon1.Visible = False
            Me.Refresh()
        Catch ex As Exception
            Error_Handler(ex, "Click on NotifyIcon")
        End Try
    End Sub


    Private Sub NotifyIcon1_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseClick
        Try
            Me.WindowState = FormWindowState.Normal
            Me.ShowInTaskbar = True
            NotifyIcon1.Visible = False
            Me.Refresh()
        Catch ex As Exception
            Error_Handler(ex, "Click on NotifyIcon")
        End Try
    End Sub
   

    Private Sub NotifyIcon1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NotifyIcon1.Click
        Try
            Me.WindowState = FormWindowState.Normal
            Me.ShowInTaskbar = True
            NotifyIcon1.Visible = False
            Me.Refresh()
        Catch ex As Exception
            Error_Handler(ex, "Click on NotifyIcon")
        End Try
    End Sub

    Private Sub Main_Screen_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        Try
            If Me.WindowState = FormWindowState.Minimized Then
                Me.ShowInTaskbar = False
                NotifyIcon1.Visible = True
                If shownminimizetip = False Then
                    NotifyIcon1.ShowBalloonTip(1)
                    shownminimizetip = True
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex, "Change Window State")
        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Try
            If CheckBox1.Checked = True Then
                NumericUpDown1.Enabled = True
                SaveSettings()
                DateTimePicker1.Enabled = False
                DateTimePicker2.Enabled = False
                DateTimePicker2.Value = Now.AddSeconds(-2)
                DateTimePicker1.Value = DateTimePicker2.Value
                DateTimePicker1.Value = DateTimePicker1.Value.AddMinutes(NumericUpDown1.Value)
            Else
                loadSettings()
                DateTimePicker1.Enabled = True
                DateTimePicker2.Enabled = True
                NumericUpDown1.Enabled = False
            End If

        Catch ex As Exception
            Error_Handler(ex, "Enable Interval based timer")
        End Try
    End Sub

    Private Sub NumericUpDown1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown1.ValueChanged
        Try
            ' DateTimePicker1.Value = DateTimePicker1.Value.AddMinutes(NumericUpDown1.Value)
            DateTimePicker1.Value = Now().AddMinutes(NumericUpDown1.Value)
            My.Settings.TimeInterval = NumericUpDown1.Value
            My.Settings.Save()
        Catch ex As Exception
            Error_Handler(ex, "Increase interval")
        End Try
    End Sub
End Class
