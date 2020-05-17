<%@Language="VBSCRIPT"%>
<% 
dim strUser, strDomain
strUser = "test_user"
strDomain = "TRENDSDI"

Set objUser = GetObject("WinNT://" & strDomain & "/" & strUser & "")

	If objUser.IsAccountLocked = 0 Then
		LockoutStatus = "Account was Not Locked"
		Response.Write "<H3>" & Server.HTMLEncode(objUser.name) & "<h3>"
		'WScript.echo(objUser.name)
		Else
		Response.Write "<H3>" & Server.HTMLEncode(objUser.name) & "<h3>"
		objUser.IsAccountLocked = 0
		objUser.SetInfo
		LockoutStatus = "Account was unlocked"
	End If

Set objUser = Nothing
Set NewPassword = Nothing
%>