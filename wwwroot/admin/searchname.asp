<%
'    Copyright 2006 Ben Norcutt and Alex Kitching
'
'    This file is part of Self Service Password Reset.
'
'    Self Service Password Reset is free software; you can redistribute it
'    and/or modify it under the terms of the GNU General Public License as
'    published by the Free Software Foundation; either version 2 of the
'    License, or (at your option) any later version.
'
'    Self Service Password Reset is distributed in the hope that it will
'    be useful, but WITHOUT ANY WARRANTY; without even the implied warranty
'    of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Self Service Password Reset; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

Dim INCLUDED
INCLUDED = True
PageTitle = "Administration"
PageMenu = True
Dim PageContent(1)

On Error Resume Next
%>
<!--#include virtual="/config.asp"-->
<!--#include file="auth.asp"-->
<!--#include virtual="/functions.asp"-->
<!--#include file="menu.asp"-->
<%
	Dim strusername, strfamilyname, Domain, arrDomain
	strgivenname = Request.Form("givenname")
	strfamilyname = Request.Form("familyname")
	If IsEmpty(strgivenname) OR IsNull(strgivenname) OR strgivenname = "" Then
	    strgivenname = "*"
	Else
	    strgivenname = "*" & LCase(strgivenname) & "*"
	End If
	If IsEmpty(strfamilyname) OR IsNull(strfamilyname) OR strfamilyname = "" Then
	    strfamilyname = "*"
	Else
	    strfamilyname = "*" & LCase(strfamilyname) & "*"
	End If

If (Not AuthAdmin) AND (Not AuthReset) Then
	PageSubTitle = "Authorised Access Only"
Else
	If AuthAdmin Then
		PageSubTitle = "Full Access Granted"
	Else If AuthReset Then
		PageSubTitle = "Reset Access Granted"
	End If
	End If

    set conn = createobject("ADODB.Connection")
    Set iAdRootDSE = GetObject("LDAP://RootDSE")
    strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
    Conn.Provider = "ADsDSOObject"
    Conn.Open "ADs Provider"

    strQueryDL = "<LDAP://" & strDefaultNamingContext & ">;(&(objectClass=person)(objectClass=user)(givenName=" & strgivenname & ")(sn=" & strfamilyname & ")(!(objectClass=computer)));distinguishedName, cn, displayName, memberOf, name, givenName, sn, homeDirectory, accountExpires,sAMAccountName, userPrincipalName;subtree"
    strQueryDL2 = "&lt;LDAP://" & strDefaultNamingContext & "&gt;;(&(objectClass=person)(objectClass=user)(givenName=" & strgivenname & ")(sn=" & strfamilyname & ")(!(objectClass=computer)));distinguishedName, cn, displayName, memberOf, name, givenName, sn, homeDirectory, accountExpires,sAMAccountName, userPrincipalName;subtree"
    set objCmd = createobject("ADODB.Command")
    objCmd.ActiveConnection = Conn
    objCmd.Properties("SearchScope") = 2 ' we want to search everything
    objCmd.Properties("Page Size") = 500 ' and we want our records in lots of 500 

    objCmd.CommandText = strQueryDL
    Set objRs = objCmd.Execute

    PageContent(1) = PageContent(1) & "<table><tr><th>Name</th><th>Year Group</th><th>Username</th><th></th></tr>"
    While Not objRS.eof
        Set objUser = GetObject("LDAP://" & objRS.Fields("distinguishedName"))
        objGroups = objUser.GetEx("memberOf")
	    For Each objGroup in objGroups
            Dim OU, pos1, pos2

            pos1 = InStr(1, objGroup, "CN=")
            pos2 = InStr(pos1+3, objGroup, "OU=")
            OU = Mid(objGroup, pos1+3, pos2-pos1-4)
		    If usersGroup = OU Then AuthResetable = True End If
		    pos1 = Nothing
		    pos2 = Nothing
		    OU = Nothing
    	Next
    	objGroups = Nothing
    	objUser = Nothing
    	If AuthAdmin OR AuthResetable Then
            PageContent(1) = PageContent(1) & "<tr><td>" & objRS.Fields("displayName") & "</td><td>" _
            & YearGroup(objRS.Fields("distinguishedName")) _
            & "</td><td>" & objRS.Fields("sAMAccountName") & "</td><td><form method=""post"" id=""search"" action=""search.asp""><input type=""hidden"" name=""username"" value=""" & objRS.Fields("sAMAccountName") & """ /><input type=""submit"" value=""Select"" /></form></td></tr>"

        End If
        AuthResetable = False
        objRS.MoveNext
    Wend
    PageContent(1) = PageContent(1) & "</table>"
End If
%>
<!--#include virtual="/template.asp"-->
