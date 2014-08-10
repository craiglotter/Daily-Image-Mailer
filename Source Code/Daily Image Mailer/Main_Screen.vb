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

    Dim DateTimePicker1_Save As String = ""
    Dim DateTimePicker2_Save As String = ""

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
                        Dim tin As FileInfo = New FileInfo(attachmentpath)
                        Dim filesize As Long = tin.Length

                        Dim filesize_string As String = ""
                        If filesize < 1024 Then
                            filesize_string = filesize & " bytes"
                        End If
                        If filesize < 1048576 And filesize >= 1024 Then
                            filesize = Math.Round(filesize / 1024, 2)
                            filesize_string = filesize & " KB"
                        End If
                        If filesize < 1073741824 And filesize >= 1048576 Then
                            filesize = Math.Round(filesize / 1048576, 2)
                            filesize_string = filesize & " MB"
                        End If
                        If filesize >= 1073741824 Then
                            filesize = Math.Round(filesize / 1073741824, 2)
                            filesize_string = filesize & " GB"
                        End If
                        msg.Body = "Your Daily Update contains the following attachment:" & vbCrLf & vbCrLf & "Name: " & tin.Name & vbCrLf & "Size: " & filesize_string & vbCrLf & vbCrLf & "Enjoy!"
                        msg.IsBodyHtml = False
                        tin = Nothing
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



    Private Sub loadSettings()
        Try
            Label2.Text = "Loading application settings..."




            Dim configfile As String = (Application.StartupPath & "\config.sav").Replace("\\", "\")
            If My.Computer.FileSystem.FileExists(configfile) Then
                Dim reader As StreamReader = New StreamReader(configfile)
                Dim lineread As String
                Dim variablevalue As String
                While reader.Peek <> -1
                    lineread = reader.ReadLine
                    If lineread.IndexOf("=") <> -1 Then

                        variablevalue = lineread.Remove(0, lineread.IndexOf("=") + 1)

                        If lineread.StartsWith("emailaddress=") Then
                            emailaddress.Text = variablevalue
                        End If
                        If lineread.StartsWith("mailserver=") Then
                            mailserver.Text = variablevalue
                        End If
                        If lineread.StartsWith("mailserverport=") Then
                            mailserverport.Text = variablevalue
                        End If
                        If lineread.StartsWith("DateTimePicker1=") Then
                            DateTimePicker1.Value = Date.Parse(variablevalue)
                            SaveSettings_Memory()
                        End If
                        If lineread.StartsWith("DateTimePicker2=") Then
                            DateTimePicker2.Value = Date.Parse(variablevalue)
                            SaveSettings_Memory()
                        End If
                        If lineread.StartsWith("NumericUpDown1=") Then
                            NumericUpDown1.Value = Integer.Parse(variablevalue)
                        End If
                        If lineread.StartsWith("CheckBox1=") Then
                            CheckBox1.Checked = Boolean.Parse(variablevalue)
                        End If
                        If lineread.StartsWith("FullErrors_Checkbox=") Then
                            FullErrors_Checkbox.Checked = Boolean.Parse(variablevalue)
                        End If

                    End If
                End While
                reader.Close()
                reader = Nothing
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



            Label2.Text = "Application Settings successfully loaded"
        Catch ex As Exception
            Error_Handler(ex, "Load Settings")
        End Try
    End Sub


    Private Sub SaveSettings()
        Try
            Label2.Text = "Saving application settings..."
            Dim configfile As String = (Application.StartupPath & "\config.sav").Replace("\\", "\")
            Dim writer As StreamWriter = New StreamWriter(configfile, False)

            If emailaddress.Text.Length > 0 Then
                writer.WriteLine("emailaddress=" & emailaddress.Text)
            End If
            If mailserver.Text.Length > 0 Then
                writer.WriteLine("mailserver=" & mailserver.Text)
            End If
            If mailserverport.Text.Length > 0 Then
                writer.WriteLine("mailserverport=" & mailserverport.Text)
            End If

            LoadSettings_Memory()

            writer.WriteLine("DateTimePicker1=" & DateTimePicker1.Value.ToString)
            writer.WriteLine("DateTimePicker2=" & DateTimePicker2.Value.ToString)

            writer.WriteLine("NumericUpDown1=" & NumericUpDown1.Value.ToString)
            writer.WriteLine("CheckBox1=" & CheckBox1.Checked.ToString)
            writer.WriteLine("FullErrors_Checkbox=" & FullErrors_Checkbox.Checked.ToString)
            

            writer.Flush()
            writer.Close()
            writer = Nothing

            Label2.Text = "Application Settings successfully saved"

        Catch ex As Exception
            Error_Handler(ex, "Save Settings")
        End Try
    End Sub






    Private Sub Main_Screen_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Try
            If CheckBox1.Checked = True Then
                LoadSettings_Memory()
            End If
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
            'If CheckBox1.Checked = False Then
            '    SaveSettings_Memory()
            'End If
        Catch ex As Exception
            Error_Handler(ex, "Change Scheduled Time")
        End Try
    End Sub

    Private Sub DateTimePicker2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker2.ValueChanged
        Try
            Label1.Text = Format(DateTimePicker2.Value, "HH:mm:ss")
            'If CheckBox1.Checked = False Then
            '    SaveSettings_Memory()
            'End If
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
                SaveSettings_Memory()
                NumericUpDown1.Enabled = True
                DateTimePicker1.Enabled = False
                DateTimePicker2.Enabled = False
                DateTimePicker2.Value = Now.AddSeconds(-2)
                DateTimePicker1.Value = DateTimePicker2.Value
                DateTimePicker1.Value = DateTimePicker1.Value.AddMinutes(NumericUpDown1.Value)
            Else

                DateTimePicker1.Enabled = True

                DateTimePicker2.Enabled = True

                LoadSettings_Memory()

                NumericUpDown1.Enabled = False
            End If

        Catch ex As Exception
            Error_Handler(ex, "Enable Interval based timer")
        End Try
    End Sub

    Private Sub NumericUpDown1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown1.ValueChanged
        Try
            If CheckBox1.Checked = True Then
                ' DateTimePicker1.Value = DateTimePicker1.Value.AddMinutes(NumericUpDown1.Value)
                DateTimePicker1.Value = Now().AddMinutes(NumericUpDown1.Value)
                'SaveSettings_Memory()
            End If

        Catch ex As Exception
            Error_Handler(ex, "Increase interval")
        End Try
    End Sub

    Private Sub SaveSettings_Memory()
        Try
            DateTimePicker1_Save = DateTimePicker1.Value.ToString
            DateTimePicker2_Save = DateTimePicker2.Value.ToString
           
        Catch ex As Exception
            Error_Handler(ex, "SaveSettings_Memory")
        End Try
    End Sub

    Private Sub LoadSettings_Memory()
        Try
            If DateTimePicker1_Save.Length > 0 Then
                DateTimePicker1.Value = Date.Parse(DateTimePicker1_Save)
            End If
            If DateTimePicker2_Save.Length > 0 Then
                DateTimePicker2.Value = Date.Parse(DateTimePicker2_Save)
            End If


        Catch ex As Exception
            Error_Handler(ex, "LoadSettings_Memory")
        End Try
    End Sub

    Private Sub Label13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label13.Click
        RunWorker()
    End Sub
End Class
