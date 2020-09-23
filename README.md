<div align="center">

## Change ANY users / user PASSWORD on the network / lan you are connected to


</div>

### Description

This code, using the windows API (NetUserChangePassword) call can change any users password on the network you are on, provided you know their original password. You need to know:

1. The machine name (ie. \\jon)

2. The username (NOT case sensitive) (ie. the_cleaner)

3. The old password (ie. password)

4. the new password (ie. password2)

Enjoy!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jon Barker](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jon-barker.md)
**Level**          |Intermediate
**User Rating**    |4.7 (33 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jon-barker-change-any-users-user-password-on-the-network-lan-you-are-connected-to__1-30170/archive/master.zip)





### Source Code

```
'PASTE THE FOLLOWING INTO ANY FORM...
'YOU MUST HAVE A COMMAND BUTTON
'NAMED 'COMMAND1'
Option Explicit
Private Declare Function NetUserChangePassword Lib "netapi32.dll" ( _
    ByVal domainname As String, ByVal Username As String, _
    ByVal OldPassword As String, ByVal NewPassword As String) As Long
Private Sub Command1_Click()
  On Error GoTo error
  Dim r As Long
  Dim sServer As String
  Dim sUser As String
  Dim sOldPass As String
  Dim sNewPass As String
  sServer = StrConv("\\jon", vbUnicode)
  sUser = StrConv("the_cleaner", vbUnicode)
  sOldPass = StrConv("password", vbUnicode)
  sNewPass = StrConv("password2", vbUnicode)
  r = NetUserChangePassword(sServer, sUser, sOldPass, sNewPass)
  If r <> 0 Then
    MsgBox "Error! Could not change password. Ensure that: " & vbCrLf & vbCrLf & _
        "o Old password was correct (Error 86)" & vbCrLf & _
        "o The server name started with '\\' (Error 1351)", vbCritical, "Error: " & r
  Else
    MsgBox "Password changed successfully!", vbExclamation, "Changed Password"
  End If
  Exit Sub
error:
  MsgBox "External error changing password: " & vbCrLf & vbCrLf & Err.Description, vbCritical, "Error: " & Err.Number
End Sub
```

