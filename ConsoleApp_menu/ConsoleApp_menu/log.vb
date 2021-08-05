Imports System.IO

Module log
    Dim path_log As String = My.Application.Info.DirectoryPath
    Dim di As New DirectoryInfo(path_log)

    Public Enum log_type
        errore = -1
        testo = 0
        info = 1
        avviso = 2
        conferma = 10
    End Enum

    Public Sub setLog(message_type As log_type, message As String)
        Dim fi As New FileInfo(path_log & "\" & Now.ToString("yyyy-MM-dd") & ".log")
        Dim fs As System.IO.StreamWriter
        fs = My.Computer.FileSystem.OpenTextFileWriter(fi.FullName, True, System.Text.Encoding.UTF8)
        Dim msg As String = Now.ToString("dd-MM-yyyyy HH:mm:ss") & vbTab & message_type.ToString & vbTab & message
        fs.WriteLine(msg)
        fs.Close()
    End Sub

    Public Function ByteUnitConvert(ByVal ByteCount As Integer) As String
dim x = ""
        Dim SetBytes As String = ""
        If ByteCount >= 1073741824 Then
            SetBytes = String.Format(ByteCount / 1024 / 1024 / 1024, "#0.00") & " GB"
        ElseIf ByteCount >= 1048576 Then
            SetBytes = Format(ByteCount / 1024 / 1024, "#0.00") & " MB"
        ElseIf ByteCount >= 1024 Then
            SetBytes = Format(ByteCount / 1024, "#0.00") & " KB"
        ElseIf ByteCount < 1024 Then
            SetBytes = Fix(ByteCount) & " Bytes"
        End If

        Return SetBytes
    End Function

End Module
