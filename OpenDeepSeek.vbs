' ============================================================
' DeepSeek 纯后台启动器（只启动服务，不打开浏览器）
' 适用场景：服务已在后台常驻，用户习惯从浏览器收藏夹访问
' ============================================================

Set WshShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

' ----- 配置区（只改这一行） -----
WorkPath = "D:\dsh"   ' 改成你的工作目录
' --------------------------------

' 确保工作目录存在
If Not FSO.FolderExists(WorkPath) Then
    FSO.CreateFolder(WorkPath)
End If

' 切换到工作目录
WshShell.CurrentDirectory = WorkPath

' 后台静默启动 dsh web（0 代表不显示命令行窗口）
WshShell.Run "%comspec% /c dsh web", 0, False

' 注意：这里没有 Sleep，也没有打开浏览器。
' 服务会在后台启动，用户自行通过浏览器收藏夹访问 http://127.0.0.1:3080