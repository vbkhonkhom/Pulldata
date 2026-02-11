Imports System.Net
Imports System.Text
Imports System.IO
Imports System.Data.SqlClient

Public Class Form1
    Public StrInputFolder As String = "C:\New folder"
    Public StrBackupFolder As String = "C:\MachineData\Backup"
    Public StrCDIR As String = System.IO.Directory.GetCurrentDirectory
    Public OKFlag As Boolean
    Public StopFlag As Boolean
    Dim MaxID As Integer = 1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not Directory.Exists(StrInputFolder) Then Directory.CreateDirectory(StrInputFolder)
        If Not Directory.Exists(StrBackupFolder) Then Directory.CreateDirectory(StrBackupFolder)
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

    Private Sub Botton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim frm As New Form2
        frm.ShowDialog()
    End Sub

    ' --- ส่วนการทำงานหลัก (แก้ไขใหม่) ---
    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        BackgroundWorker1.ReportProgress(1)
        System.Threading.Thread.Sleep(1000)

        Do
            If BackgroundWorker1.CancellationPending Then Exit Do

            Dim files As New List(Of String)()

            ' วนลูปดึงรายชื่อไฟล์ทุกนามสกุลที่กำหนด

            Try
                files.AddRange(Directory.GetFiles(StrInputFolder, "*.*"))
            Catch ex As Exception
            End Try


            If files.Count > 0 Then
                For Each filePath As String In files
                    If BackgroundWorker1.CancellationPending Then Exit Do

                    Dim fName As String = Path.GetFileName(filePath)
                    BackgroundWorker1.ReportProgress(14, "Found file: " & fName)

                    Try
                        Dim lines() As String = File.ReadAllLines(filePath, Encoding.Default)
                        If lines.Length > 0 Then
                            Dim headerLine As String = lines(0)
                            Dim headers() As String = headerLine.Split(vbTab)
                            Dim productColIndex As Integer = -1
                            For i As Integer = 0 To headers.Length - 1
                                If headers(i).Trim().Equals("Product", StringComparison.OrdinalIgnoreCase) Then
                                    productColIndex = i
                                End If
                            Next

                            If productColIndex <> -1 Then
                                Dim productData As New Dictionary(Of String, StringBuilder)
                                For r As Integer = 1 To lines.Length - 1
                                    Dim currentLine As String = lines(r)
                                    If String.IsNullOrWhiteSpace(currentLine) Then Continue For
                                    Dim cols() As String = currentLine.Split(vbTab)
                                    If cols.Length > productColIndex Then
                                        Dim productName As String = cols(productColIndex).Trim()
                                        If productName = "" Then productName = "Unknown"
                                        For Each c As Char In Path.GetInvalidFileNameChars()
                                            productName = productName.Replace(c, "_"c)
                                        Next
                                        If Not productData.ContainsKey(productName) Then
                                            productData.Add(productName, New StringBuilder())
                                            productData(productName).AppendLine(headerLine)
                                        End If
                                        productData(productName).AppendLine(currentLine)
                                    End If
                                Next
                                For Each kvp As KeyValuePair(Of String, StringBuilder) In productData
                                    Dim prodName As String = kvp.Key
                                    Dim content As String = kvp.Value.ToString()
                                    Dim newFileName As String = prodName & ".txt"
                                    Dim destPath As String = Path.Combine(StrBackupFolder, newFileName)
                                    File.WriteAllText(destPath, content, Encoding.Default)
                                    BackgroundWorker1.ReportProgress(19, "Saved Group: " & newFileName)
                                Next
                                Dim backupOriginal As String = Path.Combine(StrBackupFolder, "Processed_" & fName & "_" & DateTime.Now.ToString("yyyyMMddHHmmss") & ".bak")
                                File.Move(filePath, backupOriginal)
                                BackgroundWorker1.ReportProgress(19, "Moved Original File to Backup")
                            Else
                                BackgroundWorker1.ReportProgress(19, "Skipped: Column 'Product' not found in " & fName)
                            End If
                        End If
                    Catch ex As Exception
                        BackgroundWorker1.ReportProgress(19, "Eror: " & ex.Message)
                    End Try
                    System.Threading.Thread.Sleep(1000)
                Next
            Else
                BackgroundWorker1.ReportProgress(12, "No files in " & StrInputFolder & "...Waiting")
                System.Threading.Thread.Sleep(2000)
            End If
        Loop
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        If e.ProgressPercentage = 1 Then Form3.Label1.Text = "Processing..."
        If e.UserState IsNot Nothing Then Form3.TextBox1.AppendText(e.UserState.ToString() & vbCrLf)
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Form3.Close()
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
        If Not Directory.Exists(StrInputFolder) Then
            MsgBox("not found" & StrInputFolder)
            Exit Sub
        End If
        MsgBox("start")
        StopFlag = False
        OKFlag = True
        Me.Hide()
        Form3.Show()
        Form3.TopMost = True
        If Not BackgroundWorker1.IsBusy Then
            BackgroundWorker1.WorkerReportsProgress = True
            BackgroundWorker1.WorkerSupportsCancellation = True
            BackgroundWorker1.RunWorkerAsync()
        End If
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