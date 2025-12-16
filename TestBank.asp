<html>
<head>
<style>
 td.whitetext {color:FFFFFF}
</style>


</head>
<body>

<p>
<%
Function Tablize(row)
	Tablize = Replace(Replace(row,chr(9)&chr(9),chr(9) & "&nbsp" & chr(9)),chr(9),"</td><td>")
End Function

Dim SourceURL
	SourceURL="Samples/"

Dim strQuery  ' Search string

' You could just as easily read this from some sort of input,
' but I don't need you guys roaming around our server so
' I've hard coded it to a directory I set up to illustrate
' the sample.
' NOTE: As currently implemented, this needs to end with the /
strPath = "Samples/"

' Retrieve the search string.  If empty it will return all files.
strQuery = Request.QueryString("query")

' Show our search form and a few links to some sample searches
%>
<form action="<%= Request.ServerVariables("URL") %>" method="get">
Find files whose descriptions contain:
<input type="text" name="query" value="<%= strQuery %>" />
<input type="submit" value="Find Files" />
</form>

<p>
Some sample queries:
<a href="<%= Request.ServerVariables("URL") %>">Show All Files</a>.
<a href="<%= Request.ServerVariables("URL") %>?query=Math120">Math120</a>,
<a href="<%= Request.ServerVariables("URL") %>?query=Math140">Math140</a>,
<a href="<%= Request.ServerVariables("URL") %>?query=Math160">Math160</a>,
<a href="<%= Request.ServerVariables("URL") %>?query=Math310">Math310</a>,
<a href="<%= Request.ServerVariables("URL") %>?query=Math440">Math440</a>,
</p>
<%
Dim objFSO, objTextFile

' Create an instance of the the File System Object and assign it to objFSO.
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Open the file
Set objTextFile = objFSO.OpenTextFile(Server.MapPath("catalog.txt"))
%>
<table  border="5" bordercolor="green" cellspacing="0" cellpadding="1">
<tr bgcolor="#006600">
<td align=center class="whitetext">Filename</td>
<td align=center class="whitetext">Course</td>
<td align=center class="whitetext">Format</td>
<td align=center class="whitetext">Term</td>
<td align=center class="whitetext">Instructor</td>
<td class="whitetext">&nbsp Description</td>
</tr>
<%
Do While Not objTextFile.AtEndOfStream
	Dim DataRow
	Dim Link
	Dim FirstTab
	DataRow=objTextFile.ReadLine
	FirstTab=InStr(1,DataRow,chr(9),vbBinaryCompare)
	Link = Left(DataRow,FirstTab-1)
	DataRow = Right(DataRow,Len(DataRow)-FirstTab)
	If InStr(1, DataRow, strQuery, vbTextCompare) <> 0 Then
		%>
		<tr><td><a href=" <% Response.Write(SourceURL & Link) %> "> <% Response.Write(Link) %> </a></td><td>
		<%
		Response.Write(Tablize(DataRow) & "</td></tr>")
	End If
Loop

' Close the file.
objTextFile.Close

' Release reference to the text file.
Set objTextFile = Nothing

' Release reference to the File System Object.
Set objFSO = Nothing
%>
</table>

</ul>
<p>
<a href="../">Return to Course Webpage</a>
</body>
</html>
