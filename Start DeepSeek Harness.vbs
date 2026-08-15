' ============================================================
' DeepSeek 完整启动器（启动服务 + 打开网页）
' ============================================================

' ----- ★★★ 配置区（修改这里即可） ★★★ -----
WorkPath = "D:\dsh"                    ' 工作目录
Const ENABLE_AUTO_START = False        ' 开机自启（建议手动建快捷方式）
Const STARTUP_DELAY_MS = 3000          ' ★ 启动延迟毫秒数（3000=3秒）
' -------------------------------------------

Set WshShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

' --- 自动维护开机自启（如开启） ---
ScriptPath = WScript.ScriptFullName
ScriptName = WScript.ScriptName
StartupFolder = WshShell.SpecialFolders("Startup")
StartupLink = StartupFolder & "\" & ScriptName

If ENABLE_AUTO_START Then
    If Not FSO.FileExists(StartupLink) Then
        FSO.CopyFile ScriptPath, StartupLink, True
    End If
Else
    If FSO.FileExists(StartupLink) Then
        FSO.DeleteFile StartupLink
    End If
End If

' --- 启动服务 ---
If Not FSO.FolderExists(WorkPath) Then
    FSO.CreateFolder(WorkPath)
End If
WshShell.CurrentDirectory = WorkPath
WshShell.Run "%comspec% /c dsh web", 0, False

' --- 轮询等待服务就绪（替代固定 sleep）：每 200ms 探测一次，就绪立刻打开 ---
Set Poll = CreateObject("MSXML2.XMLHTTP")
Do While True
    Poll.Open "GET", "http://127.0.0.1:3080", False
    On Error Resume Next
    Poll.Send
    If Err.Number = 0 Then Exit Do
    Err.Clear
    On Error GoTo 0
    WScript.Sleep 200
Loop
Set Poll = Nothing

' --- 打开网页 ---
WshShell.Run "http://127.0.0.1:3080", 1, False