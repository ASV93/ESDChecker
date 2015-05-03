'============================================================================
'
'    ESDChecker
'    Copyright (C) 2015 Visual Software Corporation
'
'    Author: ASV93
'    File: Form1.vb
'
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along
'    with this program; if not, write to the Free Software Foundation, Inc.,
'    51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
'
'============================================================================

Imports System.Net
Imports System.IO
Imports System.Threading
Imports Newtonsoft.Json
Public Class Form1

    Dim BuildsLoaded As Integer = 0
    Dim BuildsFound As Integer = 0
    Dim BuildsTested As Integer = 0
    Dim ESDLimit As Integer = 10
    Dim AutoChecking As Boolean = False
    Dim chkbuild As Thread
    Dim chkweb As Thread
    Dim label6 As String 'full esd name
    Dim label3 As String 'server folder
    Dim label4 As String 'year
    Dim label5 As String 'month

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Control.CheckForIllegalCrossThreadCalls = False
        Me.Text = "Visual Software Corporation ESD Checker v" & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & My.Application.Info.Version.Build
        MetroTextBox3.Text = Now.Year
        MetroTextBox4.Text = Now.Month
        If MetroTextBox4.Text.Length = 1 Then
            MetroTextBox4.Text = "0" & MetroTextBox4.Text
        End If
        MetroTabControl1.SelectTab(0)
        Try
            ESDLimit = 1
            chkbuild = New Thread(AddressOf CheckForESDOnLine)
            chkbuild.Start()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub MetroButton6_Click(sender As Object, e As EventArgs) Handles MetroButton6.Click
        Dim IsABuild As Boolean = True
        If MetroTextBox1.Text = "" OrElse MetroTextBox2.Text = "" Then
            MessageBox.Show("You can't leave build/sha-1 field(s) empty", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If MetroTextBox1.Text.Contains("fre") OrElse MetroTextBox1.Text.Contains("chk") OrElse MetroTextBox1.Text.Contains("fbl") OrElse MetroTextBox1.Text.Contains("vol") Then
        Else
            'Not a build
            IsABuild = False
            If MetroCheckBox1.Checked = False Then
                MessageBox.Show("The string you typed in ESD Name field doesn't seem like a build, you will need to tick Build Date checkbox if you want to test it", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
        End If
        Dim LoadBuild As Boolean = True
        Dim chardir As String = "both"
        Dim FileExt = MetroTextBox1.Text.Substring(MetroTextBox1.Text.Length - 4, 4)
        label6 = MetroTextBox1.Text.Replace(FileExt, "") 'Remove ext
        label3 = "c"   'Set directory c/d
        Dim GetDate As String() = label6.Split(".")
        If MetroCheckBox1.Checked = True Then
            label4 = MetroTextBox3.Text
            label5 = MetroTextBox4.Text
        Else
            If label6.Contains("_ESD") Then
                label4 = "20" & GetDate(3).Substring(0, 2) 'Get year
                label5 = GetDate(3).Substring(2, 2) 'Get month	
            Else
                label4 = "20" & GetDate(2).Substring(0, 2) 'Get year
                label5 = GetDate(2).Substring(2, 2) 'Get month	
            End If
        End If
        If MetroCheckBox2.Checked = True Then
            Select Case label4
                Case "2015"
                    Select Case label5
                        Case "01"
                            chardir = "both"
                        Case Else
                            chardir = "both"
                    End Select
                Case Else '2014
                    Select Case label5
                        Case "08"
                            chardir = "both"
                        Case "09"
                            chardir = "both"
                        Case "11"
                            chardir = "d"
                        Case "12"
                            chardir = "both"
                        Case Else
                            If MetroToggle3.Checked = True Then
                                LoadBuild = False
                            Else
                                chardir = "both"
                            End If

                            If AutoChecking = False Then
                                MessageBox.Show("Sorry your fucking build is not going to be in any public server, don't even try to test that, but if you insist...")
                                chardir = "both"
                                LoadBuild = True
                            End If
                    End Select
            End Select
        Else
            chardir = "both"
        End If


        If IsABuild = True Then
            If MetroToggle1.Checked = True Then
                'Get Language
                Dim SplitBString As String() = MetroTextBox1.Text.Split("_")
                'Check Language
                Dim lvItem As ListViewItem = ListView2.FindItemWithText(SplitBString(SplitBString.GetLength(0) - 1).Replace(FileExt, ""), True, 0, True)
                If (lvItem IsNot Nothing) Then

                Else
                    If AutoChecking = True Then
                        LoadBuild = False
                    Else
                        MessageBox.Show("Warning: You chose to skip that language in the options tab...")
                    End If

                End If
            End If

            If MetroToggle2.Checked = True Then
                If MetroTextBox1.Text.Contains("VOL") Then
                    If AutoChecking = True Then
                        LoadBuild = False
                    Else
                        MessageBox.Show("Warning: You chose to skip VOLUME builds in the options tab...")
                    End If
                End If
            End If
        End If

        If LoadBuild = True Then

            Dim ServerURL As String
            If MetroCheckBox2.Checked = True Then
                ServerURL = "http://b1.download.windowsupdate.com/"
                If chardir = "both" Then
                    ListView1.Items.Add(ServerURL & "d" & "/updt/" & label4 & "/" & label5 & "/" & label6 & "_" & MetroTextBox2.Text & FileExt)
                    ListView1.Items.Add(ServerURL & "c" & "/updt/" & label4 & "/" & label5 & "/" & label6 & "_" & MetroTextBox2.Text & FileExt)
                    BuildsLoaded += 2
                Else
                    BuildsLoaded += 1
                    ListView1.Items.Add(ServerURL & chardir & "/updt/" & label4 & "/" & label5 & "/" & label6 & "_" & MetroTextBox2.Text & FileExt)
                End If
            End If
            If MetroCheckBox3.Checked = True Then
                ServerURL = "http://sh.dl.ws.microsoft.com/dl/content/"
                If chardir = "both" Then
                    ListView1.Items.Add(ServerURL & "d" & "/updt/" & label4 & "/" & label5 & "/" & label6 & "_" & MetroTextBox2.Text & FileExt)
                    ListView1.Items.Add(ServerURL & "c" & "/updt/" & label4 & "/" & label5 & "/" & label6 & "_" & MetroTextBox2.Text & FileExt)
                    BuildsLoaded += 2
                Else
                    BuildsLoaded += 1
                    ListView1.Items.Add(ServerURL & chardir & "/updt/" & label4 & "/" & label5 & "/" & label6 & "_" & MetroTextBox2.Text & FileExt)
                End If
            End If
            If MetroCheckBox4.Checked = True Then
                ServerURL = "http://vg.dl.ws.microsoft.com/dl/content/"
                If chardir = "both" Then
                    ListView1.Items.Add(ServerURL & "d" & "/updt/" & label4 & "/" & label5 & "/" & label6 & "_" & MetroTextBox2.Text & FileExt)
                    ListView1.Items.Add(ServerURL & "c" & "/updt/" & label4 & "/" & label5 & "/" & label6 & "_" & MetroTextBox2.Text & FileExt)
                    BuildsLoaded += 2
                Else
                    BuildsLoaded += 1
                    ListView1.Items.Add(ServerURL & chardir & "/updt/" & label4 & "/" & label5 & "/" & label6 & "_" & MetroTextBox2.Text & FileExt)
                End If
            End If
        End If
    End Sub

    Private Sub MetroButton2_Click(sender As Object, e As EventArgs) Handles MetroButton2.Click
        BuildsTested = 0
        MetroButton2.Enabled = False
        MetroButton7.Enabled = True
        PictureBox1.Image = PictureBox2.Image
        chkbuild = New Thread(AddressOf Me.CheckForValidLinksInBG)
        chkbuild.Start()
    End Sub

    Private Sub ClearToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearToolStripMenuItem.Click
        ListView1.Items.Clear()
        BuildsFound = 0
        BuildsLoaded = 0
        BuildsTested = 0
        PictureBox1.Image = PictureBox3.Image
    End Sub

    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        For Each item As ListViewItem In ListView1.SelectedItems
            Clipboard.SetText(item.Text)
        Next
    End Sub

    Private Sub MetroButton5_Click(sender As Object, e As EventArgs) Handles MetroButton5.Click
        If OpenFileDialog1.ShowDialog() = vbOK Then
            Dim currentline As Integer = 0
            For Each line In IO.File.ReadAllLines(OpenFileDialog1.FileName)
                Dim linecontent As String()
                linecontent = line.Split("	")
                If currentline = 0 Then
                    'Table Headers
                Else
                    MetroTextBox1.Text = linecontent(0)
                    MetroTextBox2.Text = linecontent(3)
                    If MetroTextBox2.Text = "" Then

                    Else
                        MetroButton6_Click(sender, e)
                    End If
                End If
                currentline = currentline + 1
            Next
        End If
        MetroTextBox1.Text = ""
        MetroTextBox2.Text = ""
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        MetroLabel3.Text = BuildsTested & " / " & BuildsLoaded & " link(s) tested (" & BuildsFound & " valid link(s) found)"
        If MetroToggle1.Checked = True Then
            PictureBox5.Visible = True
        Else
            PictureBox5.Visible = False
        End If
        If MetroToggle2.Checked = True Then
            PictureBox6.Visible = True
        Else
            PictureBox6.Visible = False
        End If
        If MetroToggle3.Checked = True Then
            PictureBox7.Visible = True
        Else
            PictureBox7.Visible = False
        End If
    End Sub

    Private Sub MetroButton3_Click(sender As Object, e As EventArgs) Handles MetroButton3.Click
        Dim newlang = InputBox("Type a language code:")
        If newlang = "" Then
        Else
            ListView2.Items.Add(newlang)
        End If
    End Sub

    Private Sub CheckForValidLinksInBG()
        'GET SIZE
        Dim currentitem As Integer = 0
        For Each item As ListViewItem In ListView1.Items
            Try
                item.ForeColor = Color.White
                Dim Request As System.Net.WebRequest
                Dim Response As System.Net.WebResponse
                Dim FileSize As Long
                Request = Net.WebRequest.Create(item.Text)
                Request.Method = Net.WebRequestMethods.Http.Head
                Request.Timeout = 5000
                Response = Request.GetResponse
                FileSize = Response.ContentLength
                item.BackColor = Color.Green
                BuildsFound += 1
                ExportToLogFile(item.Text)
                Request.Abort()
            Catch ex As Exception
                If ex.Message.Contains("timed out") Then
                    item.BackColor = Color.Yellow
                ElseIf ex.Message.Contains("404")
                    item.BackColor = Color.Red
                Else
                    item.BackColor = Color.Orange
                End If
            End Try
            ListView1.Items(currentitem).EnsureVisible()
            currentitem += 1
            BuildsTested += 1
        Next
        MetroButton7.Enabled = False
        MetroButton2.Enabled = True
        PictureBox1.Image = PictureBox3.Image
        If BuildsFound > 0 Then
            Process.Start("BuildsFound.log")
        End If
    End Sub

    Private Sub ExportToLogFile(ByVal BString As String)
        Dim strFile = "BuildsFound.log"
        Dim LineToWrite = "[" & DateTime.Today.ToString("dd-MMM-yyyy") & "] " & BString & vbCrLf
        File.AppendAllText(strFile, LineToWrite)
    End Sub

    Private Sub MetroCheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles MetroCheckBox1.CheckedChanged
        If MetroCheckBox1.Checked = True Then
            MetroTextBox3.ReadOnly = False
            MetroTextBox4.ReadOnly = False
        Else
            MetroTextBox3.ReadOnly = True
            MetroTextBox4.ReadOnly = True
        End If
    End Sub

    Private Sub MetroButton1_Click(sender As Object, e As EventArgs) Handles MetroButton1.Click
        If MetroRadioButton1.Checked = True Then
            AutoChecking = True
            MetroButton5_Click(sender, e)
        ElseIf MetroRadioButton2.Checked = True
            AutoChecking = False
            MetroButton6_Click(sender, e)
        Else
            AutoChecking = True
            MetroButton8_Click(sender, e)
        End If
    End Sub

    Private Sub MetroButton7_Click(sender As Object, e As EventArgs) Handles MetroButton7.Click
        chkbuild.Abort()
        PictureBox1.Image = PictureBox4.Image
        MetroButton7.Enabled = False
        MetroButton2.Enabled = True
    End Sub

    Private Sub MetroButton4_Click(sender As Object, e As EventArgs) Handles MetroButton4.Click
        If ListView2.SelectedItems IsNot Nothing Then
            For Each i As ListViewItem In ListView2.SelectedItems
                ListView2.Items.Remove(i)
            Next
        End If
    End Sub

    Private Sub MetroButton8_Click(sender As Object, e As EventArgs) Handles MetroButton8.Click
        If ESDLimit = 1 Then
            ESDLimit = 10
        End If
        chkbuild = New Thread(AddressOf CheckForESDOnLine)
        chkbuild.Start()
    End Sub

    Private Sub CheckForESDOnLine()
        Try
            'Make JSON Request
            Dim webClient As New System.Net.WebClient
            Dim result As String = webClient.DownloadString("http://ms-vnext.net/Win10esds/api/v1/?token=d54aaff4-bbe0-4b18-aec3-2235175514cc&columns=fileName,sha1&orderBy=fileName&orderDir=DESC&limit=" & ESDLimit)
            'Remove crap from the beginning
            Dim numberofc As Integer = 0
            Dim currentchar As Integer = 1
            While numberofc = 0
                If result(currentchar) = "{" Then
                    numberofc = 1
                Else
                    currentchar += 1
                End If
            End While
            result = result.Remove(0, currentchar)
            'Get info from JSON
            Dim jsoning As String = result.Replace("[", "")
            jsoning = jsoning.Replace("]", "")
            Dim jsonsplit As String() = jsoning.Split("{")
            Dim maxindex As Integer = jsonsplit.Length
            Dim currentindex As Integer = 1
            While currentindex < maxindex
                Dim obj As JSON_result
                If currentindex = maxindex - 1 Then
                    obj = JsonConvert.DeserializeObject(Of JSON_result)("{" & jsonsplit(currentindex).Remove(jsonsplit(currentindex).Length - 1, 1))

                Else
                    obj = JsonConvert.DeserializeObject(Of JSON_result)("{" & jsonsplit(currentindex).Remove(jsonsplit(currentindex).Length - 2, 2) & "}")
                End If
                MetroTextBox1.Text = obj.FileName
                MetroTextBox2.Text = obj.SHA1
                If AutoChecking = True Then
                    MetroButton6_Click(Nothing, Nothing)
                End If
                currentindex += 1
            End While
        Catch ex As Exception
            MessageBox.Show("Critical Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub ModifyQuantityToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ModifyQuantityToolStripMenuItem.Click
        Dim newlimit As String = InputBox("Type a new limit, Default: 10")
        If newlimit = "" Then

        Else
            ESDLimit = newlimit
            MetroRadioButton3.Text = "Import the latest " & ESDLimit & " builds from the web"
        End If
    End Sub
End Class

Public Class JSON_result
    Public FileName As String
    Public SHA1 As String
End Class