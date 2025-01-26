Option Explicit  

Dim objShell, strTarget, proxyAddress, proxyEnabled, proxyExceptions  

' 设置要删除的凭证名称  
strTarget = "192.168.0.111" ' 替换为要删除的凭证名称  

' 设置代理服务器信息  
proxyAddress = "10.32.46.4:8080"  ' 替换为您的代理服务器地址和端口  
proxyEnabled = 1                   ' 1 启用，0 禁用  
proxyExceptions = "21.*"           ' 替换为您希望排除的地址，使用分号分隔  

On Error Resume Next  

' 创建WScript.Shell对象  
Set objShell = CreateObject("WScript.Shell")  

' 删除指定的Windows凭证  
Dim cmd  
cmd = "cmd /c cmdkey /delete:" & strTarget  
objShell.Run cmd, 0, True  

' 检查凭证删除是否成功  
If Err.Number = 0 Then  
    WScript.Echo "Successfully deleted the credential: " & strTarget  
Else  
    WScript.Echo "Error deleting the credential: " & strTarget  
End If  

' 更新注册表设置以配置代理  
objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", proxyEnabled, "REG_DWORD"  
objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer", proxyAddress, "REG_SZ"  
objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyOverride", proxyExceptions, "REG_SZ"  ' 添加例外设置  

' 输出确认信息  
WScript.Echo "代理设置已更新: " & proxyAddress & vbCrLf & "例外设置: " & proxyExceptions  

' 释放对象  
Set objShell = Nothing
