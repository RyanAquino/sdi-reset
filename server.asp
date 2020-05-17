<%@Language="VBScript"%>
<!-- #include file = "text.asp" -->
<%On Error goto 0%>
<%if Request.Form("cancel") <> "" then
	if Request.Form("denyifcancel") <> "" then
		Response.Status = "401 Unauthorized"
		response.ContentType = "application/json"
		response.write("{ ""Error"":""Not authorized!""}")
		Response.End
	else
		Response.Redirect(Request.QueryString)
	end if
	Response.End
end if 
%>

<%if Request.Form("new") <> Request.Form("new2") then %>
	<% 
	response.ContentType = "application/json"
	response.write("{ ""Error"":""Password does not match""}")
	Response.End
	%>
<%end if%>

<%
	On Error resume next
	dim domain, posbs, posat, username, pUser, root
	dim upn_name

	upn_name = ""

	domain = Trim(Request.Form("domain"))
	' if no domain is present we try to get the domain from the username, 
	' e.g. domainusername or praesi@ultraschallpiloten.com
	
	if domain = "" then
		posbs = Instr(1,Request.Form("acct"),"\" )
		posat = Instr(1,Request.Form("acct"),"@" )
		if posbs > 0 then
			domain = Left(Request.Form("acct"),posbs-1)
			username = Right(Request.Form("acct"),len(Request.Form("acct")) - posbs)
		elseif posat > 0 then
			upn_name = Request.Form("acct")
			domain = Right(upn_name, len(upn_name) - posat)
			username = Left(upn_name, posat-1)
		else	
			username = Request.Form("acct")
			set nw = Server.CreateObject("WScript.Network")
			domain = nw.Computername
		end if 
	else
		username = Trim(Request.Form("acct"))
	end if
	' verify that the characters in the user name are valid
	if IsInvalidUsername(username) = true then
		response.ContentType = "application/json"
		response.write("{ ""Error"":""Username not valid""}")
		Response.End  
	end if
	
	' verify that the characters in the domain name are valid
	if IsInvalidDomainname(domain) = true then
		response.ContentType = "application/json"
		response.write("{ ""Error"":""Domain not valid""}") 
		Response.End 
	end if  
	
	if upn_name = "" then
		set pUser = GetObject("WinNT://" & domain & "/" & username & ",user")

		if Not IsObject(pUser) then
			set root = GetObject("WinNT:")
			set pUser = root.OpenDSObject("WinNT://" & domain & "/" & username & ",user", username, Request.Form("old"),1)
			' Response.Write "OpenDSObject call"
		end if

		if Not IsObject(pUser) then
			set pUser = Server.CreateObject("IIS.PwdChg")
			pUser.Domain = domain
			pUser.User = username
		end if
	else
		set pUser = Server.CreateObject("IIS.PwdChg")
		if Not IsObject(pUser) then
			set pUser = GetObject("WinNT://" & domain & "/" & username & ",user")
			if Not IsObject(pUser) then
				set root = GetObject("WinNT:")
				set pUser = root.OpenDSObject("WinNT://" & domain & "/" & username & ",user", username, Request.Form("old"),1)
				' Response.Write "OpenDSObject call"
			end if
		else
			pUser.Domain = domain
			pUser.User = username
			pUser.UPN = upn_name
		end if
	end if

	if Not IsObject(pUser) then
		'Response.Write "domain <> null - OpenDSObject also failed"
		if err.number = -2147024843 then
			response.ContentType = "application/json"
			response.write("{ ""Error"":""The specified domain or account did not exist""}")
			Response.End
		else 
			if err.description <> "" then
				response.ContentType = "application/json"
				response.write("{ ""Error:" & err.description & "}")
				response.end
			else
				response.ContentType = "application/json"
				Response.Write("{ ""Error"":"  & err.number & "}")
				response.end
			end if
		end if
		Response.End
	end if
	
	err.Clear
	pUser.ChangePassword Request.Form("old"), Request.Form("new")

	if err.number <> 0 and err.number <> 5 then
		if err.number = -2147024810 then
			response.ContentType = "application/json"
			response.write("{ ""Error"": ""Invalid username or password"" }")

		elseif err.number = -2147023545 then
			response.ContentType = "application/json"
			response.write("{ ""Error"":""Invalid Domain!""}")
			Response.End

		elseif err.number = -2147022651 then
		 	response.ContentType = "application/json"
			response.write("{ ""Error"": ""The password does not meet the account policy. Either the same in the past or haved been change twice today."" }")
			response.end
		elseif err.number = -2147022675 then
		 	response.ContentType = "application/json"
			response.write("{ ""Error"":""Username not valid""}")
			Response.End  
		elseif err.number = -2147022987 then
		 	response.ContentType = "application/json"
			response.write("{ ""Error"": ""Account is locked"" }")
		else
			response.ContentType = "application/json"
			Response.Write("{ ""Error"":"  & err.number & "}")
			'Response.Write("{ ""query"":""Li"", ""suggestions"":[""Liberia"",""Libyan Arab Jamahiriya"",""Liechtenstein"",""Lithuania""], ""data"":[""LR"",""LY"",""LI"",""LT""] }")
		end if
		Response.End
	else
		response.ContentType = "application/json"
		response.write("{ ""Success"":""Password changed successfully!""}")

	end if 

function IsInvalidUsername(username)
	dim re
	set re = new RegExp
	' list of invalid characters in a user name.
	re.Pattern = "[/\\""\[\]:<>\+=;,@]"
	IsInvalidUsername =  re.Test(username)
end function

function IsInvalidDomainname(domainname)
	dim re
	set re = new RegExp
	' list of invalid characters in a domain name. 
	re.Pattern = "[/\\""\[\]:<>\+=;,@!#$%^&\(\)\{\}\|~]"
	IsInvalidDomainName =  re.Test(domainname)
end function
%>
