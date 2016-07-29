Dim networkInfo
Set networkInfo = CreateObject("WScript.NetWork") 

Dim infoStr
infoStr = "User name is     " & networkInfo.UserName & vbCRLF & 
          "Computer name is " & networkInfo.ComputerName & vbCRLF & 
          "Domain Name is   " & networkInfo.UserDomain

MsgBox infoStr