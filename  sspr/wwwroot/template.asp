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
PageTitle = "Access Denied"
PageSubTitle = "This file cannot be called directly"
End If
%><?xml version="1.0" encoding="iso-8859-1" ?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<% '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN"
    '"http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
%>
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en">
	<head>
		<title>
			<%=PageTitle%>
		</title>
		<link rel="stylesheet" type="text/css" media="screen" href="../css.css" />
		<script type="text/javascript">
	<!--
<%=JavaScriptHeader%>
	//-->
		</script>
	</head>
	<body>
		<div id="pageWrapper">
			<div id="outerColumnContainer">
				<div id="innerColumnContainer">
					<div id="middleColumn">
						<div id="masthead" class="inside">
							<div><span><%=PageSubTitle%></span></div>
							<h1><%=PageTitle%></h1>
							<hr class="hide" />
						</div>
						<div id="content">
								<%
							If PageMenu Then
								Response.Write"<div class=""hnav bottomBorderOnly"">"
								Response.Write "<ul>"
								For a=0 To UBound(arrPageMenu)
									If (Not IsEmpty(arrPageMenu(a,0))) OR (Not IsEmpty(arrPageMenu(a,1))) Then
										Response.Write "<li><a href=""" & arrPageMenu(a,0) & """>" & arrPageMenu(a,1) & "</a></li>"
									End If
								Next
								Response.Write "</ul>"
								Response.Write "<hr class=""hide"" />" & vbNewLine & "</div>"
							End If
							For Each obj In PageContent
								If (Not IsEmpty(obj)) Then
									Response.Write("<div class=""main"">")
									Response.Write(obj)
									Response.Write("</div>")
								End If
							Next
						%>
						</div>
					</div>
				<br style="clear: both;" />
				</div>
			</div>
		</div>
	</body>
</html>
