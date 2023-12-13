    
on error resume Next
Const HKLM         = &H80000002   
Const HKCU         = &H80000001
Const strKeyPath   = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
Const strKeyPathwin  ="SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
Const ForReading   = 1   
Const ForAppending = 8  
  

Const FilePath     ="\\dir\\"
Set Wshell         = CreateObject("Wscript.Shell")   
Set objFSO         = CreateObject("Scripting.FileSystemobject")   

Dim WshNetwork
Set WshNetwork = WScript.CreateObject("WScript.Network")
strComputer = WshNetwork.ComputerName  


Set objShell = CreateObject("WScript.Shell")  
Set objScriptExec = objShell.Exec("ipconfig")  
strIpConfig = objScriptExec.StdOut.ReadAll  
Set objRegEx = CreateObject("VBScript.RegExp")  
objRegEx.Global = True  
objRegEx.Pattern = "(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})"  
Set colMatches = objRegEx.Execute(strIpConfig) 

If colMatches.Count > 0 Then         
		if (objFSO.FileExits(FilePath & colMatches(0).Value & "(" & WshNetwork.ComputerName & ")" &".txt")) then
		Set MyFile = objFSO.GetFile(FilePath & colMatches(0).Value & "(" & WshNetwork.ComputerName & ")" &".txt") 
		MyFile.Delete
		End If 
End If
Set textWriteFile  = objFSO.OpenTextFile(FilePath & colMatches(0).Value & "(" & WshNetwork.ComputerName & ")" &".txt",forappending,True)


Set objReg  = GetObject("winmgmts://" & strComputer & "/root/default:StdRegProv")   
objReg.EnumKey HKCU, strKeyPath,arrSubKeys    

  For Each strSubKey In arrSubKeys        
    intRet = objReg.GetStringValue(HKCU, strKeyPath & strSubKey,"DisplayName",strValue)                                                    
         If strValue <> "" And intRet = 0 And Left(strSubKey, 2) <> "KB" Then   
              textWriteFile.WriteLine(strValue)
        End If 
 Next  

Set objReglm  = GetObject("winmgmts://" & strComputer & "/root/default:StdRegProv")   
objReg.EnumKey HKLM, strKeyPath,arrSubKeys    

  For Each strSubKey In arrSubKeys        
    intRet = objReg.GetStringValue(HKLM, strKeyPath & strSubKey,"DisplayName",strValue)                                                    
        If strValue <> "" And intRet = 0 And Left(strSubKey, 2) <> "KB" Then   
              textWriteFile.WriteLine(strValue)
        End If 
 Next

Set objReg  = GetObject("winmgmts://" & strComputer & "/root/default:StdRegProv")   
objReg.EnumKey HKLM, strKeyPathwin,arrSubKeys    

  For Each strSubKey In arrSubKeys        
    intRet = objReg.GetStringValue(HKLM, strKeyPathwin & strSubKey,"DisplayName",strValue)                                                    
      If strValue <> "" And intRet = 0 And Left(strSubKey, 2) <> "KB" Then   
              textWriteFile.WriteLine(strValue)
        End If 
 Next
