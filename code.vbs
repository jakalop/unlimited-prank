Dim networkInfo
Set networkInfo = CreateObject("WScript.NetWork") 

Dim infoStr
infoStr = "User name is     " & networkInfo.UserName & vbCRLF & _
          "Computer name is " & networkInfo.ComputerName & vbCRLF & _
          "Domain Name is   " & networkInfo.UserDomain

MsgBox infoStr

