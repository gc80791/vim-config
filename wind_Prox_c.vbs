Option Explicit  

Dim objShell, strTarget, proxyAddress, proxyEnabled, proxyExceptions  

' ����Ҫɾ����ƾ֤����  
strTarget = "192.168.0.111" ' �滻ΪҪɾ����ƾ֤����  

' ���ô����������Ϣ  
proxyAddress = "10.32.46.4:8080"  ' �滻Ϊ���Ĵ����������ַ�Ͷ˿�  
proxyEnabled = 1                   ' 1 ���ã�0 ����  
proxyExceptions = "21.*"           ' �滻Ϊ��ϣ���ų��ĵ�ַ��ʹ�÷ֺŷָ�  

On Error Resume Next  

' ����WScript.Shell����  
Set objShell = CreateObject("WScript.Shell")  

' ɾ��ָ����Windowsƾ֤  
Dim cmd  
cmd = "cmd /c cmdkey /delete:" & strTarget  
objShell.Run cmd, 0, True  

' ���ƾ֤ɾ���Ƿ�ɹ�  
If Err.Number = 0 Then  
    WScript.Echo "Successfully deleted the credential: " & strTarget  
Else  
    WScript.Echo "Error deleting the credential: " & strTarget  
End If  

' ����ע������������ô���  
objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", proxyEnabled, "REG_DWORD"  
objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer", proxyAddress, "REG_SZ"  
objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyOverride", proxyExceptions, "REG_SZ"  ' �����������  

' ���ȷ����Ϣ  
WScript.Echo "���������Ѹ���: " & proxyAddress & vbCrLf & "��������: " & proxyExceptions  

' �ͷŶ���  
Set objShell = Nothing
