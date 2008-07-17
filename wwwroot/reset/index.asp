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
PageTitle = "Password Reset"
PageSubTitle = "Step 1"
PageMenu = False
Dim PageContent(1)

JavaScriptHeader = "function formValidation(form){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(notEmpty(form.username)){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "return true;" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "} else {" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "alert(""Please enter your username"");" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "document.form.username.focus();" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "return false;" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "function notEmpty(elem){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(elem.value.length == 0){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "return false;" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "} else {" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "return true;" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine

PageContent(0) = "To begin the process of resetting your password, please enter your username below"
PageContent(1) = "<form name=""form"" method=""post"" onSubmit=""return formValidation(this)"" action=""reset.asp"">Username: <input type=""text"" name=""username"" id=""username"" maxlength=""20"" /><br /><input type=""submit"" value=""Submit"" /></form>"
%>
<!--#include virtual="/template.asp"-->