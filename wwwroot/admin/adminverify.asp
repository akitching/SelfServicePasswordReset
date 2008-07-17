<%
Dim INCLUDED
INCLUDED = True
PageTitle = "Administration"
PageMenu = True
Dim PageContent(1)
%>
<!--#include virtual="/config.asp"-->
<!--#include file="auth.asp"-->
<!--#include virtual="/functions.asp"-->
<!--#include file="menu.asp"-->
<%
	Dim strusername, Domain, arrDomain
	strusername = Request.Form("username")
	strusername = LCase(strusername)
	
If (Not AuthAdmin) AND ((Not AuthReset) OR (Not AuthResetable)) Then
	PageSubTitle = "Authorised Access Only"
	If (Not AuthAdmin) AND (Not AuthReset) Then
		PageContent(0) = "<h3>You are not authorised to access this service.<br /><a href=""http://" & homepage & """>Leave</a>.</h3>"
	Else If Not AuthResetable Then
		PageContent(0) = "<h3>You are not authorised to access this user.<br /><a href=""index.asp"">Return to Administration</a>.</h3>"
	End If
	End If
Else
	If AuthAdmin Then
		PageSubTitle = "Full Access Granted"
	Else If AuthReset Then
		PageSubTitle = "Reset Access Granted"
	End If
	End If
	
	'Create an ADO connection object
	Set adoCon = Server.CreateObject("ADODB.Connection")
	'Set an active connection to the Connection object using a DSN-less connection
	adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & DataBase & ";uid=" & DBUID & ";pwd=" & DBPWD & ";"

	'Create an ADO recordset object
	Set rsReset = Server.CreateObject("ADODB.Recordset")
	'Initialise the strSQL variable with an SQL statement to query the database
	'strSQL = "SELECT tblmain.username, tblmain.registered FROM tblmain;"
	strSQL = "SELECT tblmain.* FROM tblmain where username='" & strusername & "';"
	'Open the recordset with the SQL query 
	rsReset.Open strSQL, adoCon
	
	'Loop through the recordset
	'Do While Not rsReset.EOF
	If Not rsReset.EOF Then
		If AuthAdmin OR resetGroupSemiSecret Then
			If rsReset("answer3") = LCase(Request.Form("answer")) Then
				PageContent(0) = PageContent(0) & "User Verified" & vbNewLine
			Else
				PageContent(0) = PageContent(0) & "User Not Verified" & vbNewLine
			End If
		End If
	Else
	PageContent(0) = PageContent(0) & "User not registered in database." & vbNewLine
	End If
	
End If
%>
<!--#include virtual="/template.asp"-->