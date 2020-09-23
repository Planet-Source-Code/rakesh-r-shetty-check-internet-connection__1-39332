<div align="center">

## check internet connection


</div>

### Description

checks the internet connection
 
### More Info
 
I have tested this code on windows 2000 and XP.

Dont know whether it works for windows 9x.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rakesh R\. Shetty](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rakesh-r-shetty.md)
**Level**          |Beginner
**User Rating**    |5.0 (35 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rakesh-r-shetty-check-internet-connection__1-39332/archive/master.zip)





### Source Code

```
Private Const FLAG_ICC_FORCE_CONNECTION = 1
Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Boolean
Private Sub Command1_Click()
Dim abc As String
 abc = InternetCheckConnection("http://www.microsoft.com", FLAG_ICC_FORCE_CONNECTION, 0)
 If abc = "True" Then
  MsgBox "Connected to the internet."
 Else
  MsgBox "Not connected to the internet."
 End If
End Sub
```

