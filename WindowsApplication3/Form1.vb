'Imports System.Net
Imports System.Text
Imports System.IO
'Imports System.Data.SqlClient
'Imports System.ComponentModel

Public Class Form1
    Public StrInputFolder As String = "C:\Users\admin\Documents\Peelforce Data report"
    Public StrBackupFolder As String = "C:\Users\admin\Documents\Data"
    Public StrCDIR As String = System.IO.Directory.GetCurrentDirectory
    Public OKFlag As Boolean
    Public StopFlag As Boolean
    Private lastReadPosition As Long = 0
    Private posFile As String = Path.Combine(Application.StartupPath, "last_pos.txt")

    Private Sub LoadLastPosition()
        If File.Exists(posFile) Then
            Long.TryParse(File.ReadAllText(posFile), lastReadPosition)
        End If
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error Resume Next
        LoadLastPosition()
        ReadConfigData()
        If Not Directory.Exists(StrInputFolder) Then Directory.CreateDirectory(StrInputFolder)
        If Not Directory.Exists(StrBackupFolder) Then Directory.CreateDirectory(StrBackupFolder)



        NotifyIcon1.Icon = Me.Icon
        NotifyIcon1.Text = "SPC Data Separator Service"
        NotifyIcon1.Visible = True

        Me.WindowState = FormWindowState.Minimized
        Me.ShowInTaskbar = False
        Me.Visible = False

        If Not BackgroundWorker1.IsBusy Then
            BackgroundWorker1.RunWorkerAsync()
        End If
    End Sub
    Private Sub StartService()
        On Error Resume Next
        If Not BackgroundWorker1.IsBusy Then
            StopFlag = False
            OKFlag = True
            Me.WindowState = FormWindowState.Minimized
            Me.Hide()
            NotifyIcon1.Visible = True
            BackgroundWorker1.RunWorkerAsync()
        End If
    End Sub


    Private Sub Botton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim frm As New Form2
        frm.ShowDialog()
    End Sub

    ' --- ส่วนการทำงานหลัก (แก้ไขใหม่) ---
    Private fileOffsets As New Dictionary(Of String, Long)
    Private waitCount As Integer = 0
    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Do
            If BackgroundWorker1.CancellationPending Then Exit Do
            Try
                Dim filePath As String = Path.Combine(StrInputFolder, "testdata")
                If File.Exists(filePath) Then
                    Dim fi As New FileInfo(filePath)
                    If fi.Length > lastReadPosition Then
                        ProcessIncrementalUpdate(filePath)
                        waitCount = 0
                    ElseIf fi.Length < lastReadPosition Then
                        lastReadPosition = 0
                        ProcessIncrementalUpdate(filePath)
                    End If
                    waitCount += 1
                    If waitCount >= 5 Then
                        BackgroundWorker1.ReportProgress(12, "wait..")
                        waitCount = 0
                    End If
                Else
                    BackgroundWorker1.ReportProgress(12, "Waiting for new logs...")
                End If
            Catch ex As Exception
            End Try
            System.Threading.Thread.Sleep(1000)
        Loop
    End Sub
    Private Sub ProcessIncrementalUpdate(ByVal filePath As String)
        Try
            Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                If lastReadPosition > fs.Length Then
                    lastReadPosition = 0
                    Exit Sub
                End If
                fs.Seek(lastReadPosition, SeekOrigin.Begin)
                Using sr As New StreamReader(fs, Encoding.Default)
                    Dim headerLine As String = ""

                    Using fsHeader As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                        Using srHeader As New StreamReader(fsHeader, Encoding.Default)
                            headerLine = srHeader.ReadLine()
                        End Using
                    End Using

                    If lastReadPosition = 0 Then sr.ReadLine()

                    Dim count As Integer = 0

                    'Dim productsFound As New List(Of String)
                    While Not sr.EndOfStream
                        Dim currentLine As String = sr.ReadLine()
                        If String.IsNullOrWhiteSpace(currentLine) Then Continue While
                        Dim cols() As String = currentLine.Split(ControlChars.Tab)
                        If currentLine.Contains("Lot Number") Then Continue While
                        If cols.Length > 8 Then
                            Dim finalProductName As String = cols(8).Trim()
                            If String.IsNullOrEmpty(finalProductName) Then finalProductName = "Unknow"

                            For Each c As Char In Path.GetInvalidFileNameChars()
                                finalProductName = finalProductName.Replace(c, "_"c)
                            Next
                            Dim targetFile As String = Path.Combine(StrBackupFolder, finalProductName & ".txt")
                            Using sw As New StreamWriter(targetFile, True, Encoding.Default)
                                If New FileInfo(targetFile).Length = 0 Then sw.WriteLine(headerLine)
                                sw.WriteLine(currentLine)
                            End Using
                            count += 1
                            'If Not productsFound.Contains(productName) Then productsFound.Add(productName)
                        End If
                    End While
                    lastReadPosition = fs.Position
                    File.WriteAllText(posFile, lastReadPosition.ToString())
                    If count > 0 Then
                        BackgroundWorker1.ReportProgress(14, "แยก Product สำเร็จ: " & count)
                    End If
                End Using
            End Using
        Catch ex As Exception
            BackgroundWorker1.ReportProgress(19, "Error " & ex.Message)
        End Try
    End Sub
    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            Me.Hide()
            NotifyIcon1.Visible = True
        End If
    End Sub
    Private Sub NotifyIcon1_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        Me.Show()
        Me.WindowState = FormWindowState.Normal
        Me.ShowInTaskbar = True
        Me.Visible = True
    End Sub
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If Not StopFlag Then
            e.Cancel = True
            Me.WindowState = FormWindowState.Minimized
            Me.Hide()
            NotifyIcon1.Visible = True
            NotifyIcon1.ShowBalloonTip(2000, "Service running", "Dev", ToolTipIcon.Info)
        End If
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        On Error Resume Next
        If e.UserState IsNot Nothing Then
            Form3.TextBox1.AppendText("[" & Now.ToString("T") & "]" & e.UserState.ToString() & vbCrLf)
            Form3.TextBox1.SelectionStart = Form3.TextBox1.TextLength
            Form3.TextBox1.ScrollToCaret()
        End If
    End Sub
    Private Sub ReadConfigData()
        Dim strFileName As String = StrCDIR & "\Config.csv"
        If File.Exists(strFileName) Then
            Dim sr As New StreamReader(strFileName, Encoding.Default)
            Dim line As String
            Do Until sr.Peek() = -1
                line = sr.ReadLine()
                Dim temp() As String = Split(line, ",")
                If temp.Length >= 2 Then
                    If temp(0) = "InputFolder" Then StrInputFolder = temp(1)
                End If
            Loop
            sr.Close()
        End If
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        'Form3.Close()
        Me.Show()
    End Sub

    Private Function Read_File_Spy(ByVal filePath As String, ByVal startID As Integer) As String(,)
        Dim allText As String = File.ReadAllText(filePath, Encoding.Default)
        Dim previewText As String = allText
        If previewText.Length > 200 Then previewText = previewText.Substring(0, 200) & "..."
        Dim match As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(allText, "[0-9]+\.[0-9]+")
        Dim foundValue As String = ""
        If match.Success Then
            foundValue = match.Value
        Else
            match = System.Text.RegularExpressions.Regex.Match(allText, "[0-9]+")
            If match.Success Then
                foundValue = match.Value
                If Val(foundValue) > 2000 Then
                    foundValue = ""
                End If
            End If
        End If
        If foundValue = "" Then
            MsgBox("โปรแกรมอ่านไฟล์แล้ว แต่หาตัวเลขไม่เจอ" & vbCrLf &
                   "-------------------------" & vbCrLf &
                   "สิ่งที่โปรแกรมเห็นคือ: " & vbCrLf &
                   "[" & previewText & "]" & vbCrLf &
                   "สาเหตุที่เป็นไปได้ ไฟล์เปล่า หรืออื่นๆ")
            Return Nothing
        End If
        Dim _DataBuf(0, 250) As String
        For c As Integer = 0 To 125 : _DataBuf(0, c) = "Null" : Next
        _DataBuf(0, 0) = startID.ToString()
        _DataBuf(0, 2) = "Machine_01"
        _DataBuf(0, 14) = Path.GetFileNameWithoutExtension(filePath)
        _DataBuf(0, 26) = foundValue
        Return _DataBuf
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not BackgroundWorker1.IsBusy Then
            BackgroundWorker1.RunWorkerAsync()
        End If
        Form3.Show()
        Form3.TextBox1.AppendText("[" & Now.ToString("T") & "] บังคับตรวจสอบข้อมูลด้วยตนเอง..." & vbCrLf)
    End Sub

    ' --- ฟังก์ชันใหม่: สำหรับแปลง Excel เป็น CSV ---
    Private Sub ConvertExcelToCsv(ByVal sourceExcelPath As String, ByVal destCsvPath As String)
        Dim excelApp As Object = Nothing
        Dim wb As Object = Nothing
        Try
            ' สร้าง Object Excel (Late Binding ไม่ต้อง Add Reference)
            excelApp = CreateObject("Excel.Application")
            excelApp.Visible = False
            excelApp.DisplayAlerts = False

            ' เปิดไฟล์ Excel
            wb = excelApp.Workbooks.Open(sourceExcelPath)

            ' บันทึกเป็น CSV (FileFormat 6 = xlCSV)
            wb.SaveAs(Filename:=destCsvPath, FileFormat:=6)
            wb.Close(SaveChanges:=False)
        Catch ex As Exception
            Throw New Exception("Excel Convert Error: " & ex.Message)
        Finally
            ' คืนค่าหน่วยความจำและปิด Excel ให้สนิท
            If wb IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(wb)
            If excelApp IsNot Nothing Then
                excelApp.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
            End If
            wb = Nothing
            excelApp = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

End Class