<div align="center">

## Authenticate against Windows NT/2000 Domain


</div>

### Description

This code will take the users username and password from a form and use them to authenticate them against a Windows NT/2000 domain. Unlike other examples, you do not need to switch on 'Basic' or 'Integrated Windows' permissions for the webite on IIS. You can leave the setting as Anonymous Access.
 
### More Info
 
Form inputs

Make sure that the site being protected is set to Anonymous Authentication and that the users using the site all have Windows accounts.

Whether or not the users has been authenticated

This is highly insecure over the Internet. I recommend you use SSL to protect user details.

Also, only protects ASP pages.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[MarkParter](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/markparter.md)
**Level**          |Intermediate
**User Rating**    |4.4 (53 globes from 12 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Security](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/security__4-14.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/markparter-authenticate-against-windows-nt-2000-domain__4-7814/archive/master.zip)





### Source Code

```
'Place the following in your ASP page handling the server-side authetication.
'************************************************
<% Dim objADSI, strUsername, strPassword, strDomain
strUsername = Trim(Request.Form("txtUsername"))
strPassword = Trim(Request.Form("txtPassword"))
strDomain = "Intranet"
'you can easily change this to retrieve the domain from a form aswell
Set objADSI = GetObject("WinNT://" & strDomain)
Dim strADsNamespace
Dim objADSINamespace
strADsNamespace = Left("WinNT://" & strDomain, InStr("WinNT://" & strDomain, ":"))
Set objADSINamespace = GetObject(strADsNamespace)
Set objADSI = objADSINamespace.OpenDSObject("WinNT://" & strDomain, strDomain & "\" & strUsername, strPassword, 0)
' If there's no error then the user has been authenticated!
If Err.Number <> 0 Then 'authentication failed
  'code here for failed authentication
  Session("authenticated") = False
Else
  'code here for authentication success
  Session("authenticated") = True
End If
Set objADSINamespace = Nothing
Set objADSI = Nothing
Set strUsername = Nothing
Set strPassword = Nothing
Set strDomain = Nothing
Set strADsNamespace = Nothing %>
'***********************************************
At the top of all your protected ASP pages place the following:
<!-- #INCLUDE file="check.asp" -->
Make sure you check the path to the file, if necessary make it an absolute include, i.e. <!-- #INCLUDE file="http://www.yoursite.co.uk/check.asp" -->
'************************************************
create a file called check.asp, in it place the following code:
<% If Session("authenticated") <> True Then
   Session.Abandon 'clear any session variables
   Response.Redirect "login.asp" 'kick them back to the login page
End If %>
```

