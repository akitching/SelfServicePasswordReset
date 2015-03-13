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

' Automatically log administrative password resets into a GLPI helpdesk.
INCLUDED=True
IF Not INCLUDED Then
%>
<!--#include virtual="/template.asp"-->
<%
Else

Public Function LogInHelpdesk(ByVal username)
    On Error Resume Next
    datetime=GetDateTime(Now)

        dim cn, rs, uid, HelpdeskSQL

        set cn = CreateObject("ADODB.Connection")
        set rs = CreateObject("ADODB.Recordset")
        connectionstring = "Driver={MySQL ODBC 5.1 Driver};Server=" & HelpdeskServer & ";Database=" & HelpdeskDB & ";User=" & HelpdeskDBUser & ";Password=" & HelpdeskDBPass & ";"
        cn.connectionstring = connectionstring
        cn.open
        rs.open "SELECT ID FROM glpi_users WHERE name = '" & CurrentUser & "' LIMIT 1", cn, 3
        If Not rs.EOF Then
            uid = rs("ID")
            If HelpdeskLinkToUser = 1 Then
                ' Set user being reset as ticket's owner, include generic description in ticket content.
                set rso = CreateObject("ADODB.Recordset")
                rso.open "SELECT ID FROM glpi_users WHERE name = '" & username & "' LIMIT 1", cn, 3
                If Not rso.EOF Then
                HelpdeskSQL = "INSERT INTO `glpi_tickets` (`entities_id`, `name`, `date`, `close_date`, `solve_date` `date_mod`, `status`, `users_id`, `users_id_recipient`, `requesttypes_id`, `users_id_assign`, `items_id`, `content`, `urgency`, `impact`, `priority`, `use_email_notification`, `ticketcategories_id`, `global_validation`) VALUES ('1', 'Reset Password - "& username & "', '" & datetime & "', '" & datetime & "', '" & datetime & "', '" & datetime & "' 'closed', '" & uid & "', '" & rso("id") & "', '7', '" & uid & "', '0', 'Password reset for " & username & "', '3', '1', '2', '0', '" & HelpdeskCategoryID & "', 'none');"
                Else
                    ' User not found, fall back on non linking method.
                HelpdeskSQL = "INSERT INTO `glpi_tickets` (`entities_id`, `name`, `date`, `close_date`, `solve_date` `date_mod`, `status`, `users_id`, `users_id_recipient`, `requesttypes_id`, `users_id_assign`, `items_id`, `content`, `urgency`, `impact`, `priority`, `use_email_notification`, `ticketcategories_id`, `global_validation`) VALUES ('1', 'Reset Password - "& username & "', '" & datetime & "', '" & datetime & "', '" & datetime & "', '" & datetime & "' 'closed', '" & uid & "', '" & uid & "', '7', '" & uid & "', '0', 'Password reset for " & username & "', '3', '1', '2', '0', '" & HelpdeskCategoryID & "', 'none');"
                End If
            Else
                ' Include username of user being reset into ticket content, along with generic description.
                HelpdeskSQL = "INSERT INTO `glpi_tickets` (`entities_id`, `name`, `date`, `close_date`, `solve_date` `date_mod`, `status`, `users_id`, `users_id_recipient`, `requesttypes_id`, `users_id_assign`, `items_id`, `content`, `urgency`, `impact`, `priority`, `use_email_notification`, `ticketcategories_id`, `global_validation`) VALUES ('1', 'Reset Password - "& username & "', '" & datetime & "', '" & datetime & "', '" & datetime & "', '" & datetime & "' 'closed', '" & uid & "', '" & uid & "', '7', '" & uid & "', '0', 'Password reset for " & username & "', '3', '1', '2', '0', '" & HelpdeskCategoryID & "', 'none');"
            End If
            cn.Execute HelpdeskSQL
        End If

        cn.close

End Function

Public Function GetDateTime(varDate)
    If day(varDate) < 10 Then
        dd = "0" & day(varDate)
    Else
        dd = day(varDate)
    End If

    If month(varDate) < 10 Then
        mm = "0" & month(varDate)
    Else
        mm = month(varDate)
    End If

    If hour(varDate) < 10 Then
        ho = "0" & hour(varDate)
    Else
        ho = hour(varDate)
    End If

    If minute(varDate) < 10 Then
        mi = "0" & minute(varDate)
    Else
        mi = minute(varDate)
    End If

    If second(varDate) < 10 Then
        ss = "0" & second(varDate)
    Else
        ss = second(varDate)
    End If

    GetDateTime = year(varDate) & "-" & mm & "-" & dd & " " & ho & ":" & mi & ":" & ss
End Function

End If
%>
