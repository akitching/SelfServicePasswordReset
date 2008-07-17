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

IF Not INCLUDED Then
%>
<!--#include virtual="/template.asp"-->
<%
Else
'----------
Function RandomPW(myLength)
	'These constant are the minimum and maximum length for random
	'length passwords.  Adjust these values to your needs.
	Dim minLength, maxLength
	minLength = RandomPasswordMin
	maxLength = RandomPasswordMax
	
	Dim X, Y, strPW
	
	If myLength = 0 Then
		Randomize
		myLength = Int((maxLength * Rnd) + minLength)
	End If

	
	For X = 1 To myLength
		'Randomize the type of this character
		'Y = Int((3 * Rnd) + 1) '(1) Numeric, (2) Uppercase, (3) Lowercase
		Y = Int((2 * Rnd) + 1) '(1) Numeric, (2) Uppercase, (3) Lowercase
		
		Select Case Y
			Case 1
				'Numeric character
				Randomize
				strPW = strPW & CHR(Int((9 * Rnd) + 48))
			Case 2
				'Uppercase character
				Randomize
				strPW = strPW & CHR(Int((25 * Rnd) + 65))
			Case 3
				'Lowercase character
				Randomize
				strPW = strPW & CHR(Int((25 * Rnd) + 97))

		End Select
	Next
	
	RandomPW = strPW
End Function

Function PWList(myfile, strusername)
	
  Const FOR_READING = 1
  Dim objFSO, objTS, strContents

  Set objFSO = CreateObject("Scripting.FileSystemObject")

  If objFSO.FileExists(myfile) Then
    On Error Resume Next
    Set objTS = objFSO.OpenTextFile(myfile, FOR_READING, False, blnUNICODE)
    If Err = 0 Then
      strContents = objTS.ReadAll
      objTS.Close

      arrPwd = Split(strContents, vbNewLine)

    Else
      Response.Write("<h3>Error</h3>")
    End If
  End If

max = UBound(arrPwd) -1
  For a=0 to max
    arrUser = split(arrPwd(a), ",")
    If LCase(arrUser(0)) = strusername Then
      password = arrUser(1)
    End If
  Next
	
	PWList = password
End Function

Function RandomPWList(myfile)
	Randomize

  Const FOR_READING = 1
  Dim objFSO, objTS, strContents

  Set objFSO = CreateObject("Scripting.FileSystemObject")

  If objFSO.FileExists(myfile) Then
    On Error Resume Next
    Set objTS = objFSO.OpenTextFile(myfile, FOR_READING, False, blnUNICODE)
    If Err = 0 Then
      strContents = objTS.ReadAll
      objTS.Close

      arrPwd = Split(strContents, vbNewLine)

    Else
      Response.Write("<h3>Error</h3>")
    End If
  End If

listNumber = Int(UBound(arrPwd) * Rnd)

	RandomPWList = arrPwd(listNumber)
End Function

Function PWPromptUser

PageContent(0) = "<form name=""form"" method=""post"" onsubmit=""return formValidation(this)"" action=""reset_pass.asp"">" & vbNewLine
PageContent(0) = PageContent(0) & "<input type=""hidden"" name=""username"" maxlength=""50"" value=""" & strusername & """/>" & vbNewLine
PageContent(0) = PageContent(0) & "<input type=""hidden"" name=""answer1"" value=""" & Request.Form("answer1") & """ />" & vbNewLine
PageContent(0) = PageContent(0) & "<input type=""hidden"" name=""answer2"" value=""" & Request.Form("answer2") & """ />" & vbNewLine
PageContent(0) = PageContent(0) & "<input type=""hidden"" name=""answer3"" value=""" & Request.Form("answer3") & """/>" & vbNewLine
PageContent(0) = PageContent(0) & "<table>" & vbNewLine
PageContent(0) = PageContent(0) & "<tr>" & vbNewLine
PageContent(0) = PageContent(0) & "<td class=""r1 c1""><label for=""NewPass"">Please enter your new password</label>: </td><td class=""r1 c2""><input type=""text"" name=""NewPass"" id=""NewPass"" maxlength=""50"" /></td>" & vbNewLine
PageContent(0) = PageContent(0) & "</tr>" & vbNewLine
PageContent(0) = PageContent(0) & "</table>" & vbNewLine
PageContent(0) = PageContent(0) & "<input type=""submit"" value=""Submit"" />" & vbNewLine
PageContent(0) = PageContent(0) & "</form>" & vbNewLine

End Function

Function PWPromptAdmin

PageContent(0) = "<form name=""reset"" method=""post"" onsubmit=""return formValidation(this)"" action=""adminreset.asp"">" & vbNewLine
PageContent(0) = PageContent(0) & "<input type=""hidden"" name=""username"" maxlength=""50"" value=""" & strusername & """/>" & vbNewLine
PageContent(0) = PageContent(0) & "<table>" & vbNewLine
PageContent(0) = PageContent(0) & "<tr>" & vbNewLine
PageContent(0) = PageContent(0) & "<td class=""r1 c1""><label for=""NewPass"">Please enter your new password</label>: </td><td class=""r1 c2""><input type=""text"" name=""NewPass"" id=""NewPass"" maxlength=""50"" /></td>" & vbNewLine
PageContent(0) = PageContent(0) & "</tr>" & vbNewLine
PageContent(0) = PageContent(0) & "</table>" & vbNewLine
PageContent(0) = PageContent(0) & "<input type=""submit"" name=""resetPassword"" value=""Reset Password"" />" & vbNewLine
PageContent(0) = PageContent(0) & "</form>" & vbNewLine

End Function
'----------
End If
%>