<%
Dim strPath   ' Path of directory to search
Dim objFSO    ' FileSystemObject variable
Dim objFolder ' Folder variable
Dim objItem   ' Variable used to loop through the
              'contents of the folder

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
Find files whose names contain:
<input type="text" name="query" value="<%= strQuery %>" /><br />
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
' Create our FSO and get a handle on our folder
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(Server.MapPath(strPath))

' Show a little description line and the title row of our table
%>

<table border="5" bordercolor="green" cellspacing="0" cellpadding="2">
    <tr bgcolor="#006600">
        <td><font color="#FFFFFF"><strong>File&nbsp;Name:</strong></font></td>
        <td><font color="#FFFFFF"><strong>File&nbsp;Size:</strong></font></td>
        <td><font color="#FFFFFF"><strong>Date&nbsp;Created:</strong></font></td>
    </tr>
<%
' Now come the part where we find the files whose names match the string
' passed in part.  It's relatively straightforward.  We simply loop through
' the objFolder.Files collection and check each file's name to see if it
' matches.  Note that I'm not going to look for SubFolders that match our
' search string.  You can easily include them if you like.  Just use
' objFolder.SubFolders.  The syntax is the same as it is for files.
For Each objItem In objFolder.Files
    ' I'm using a case insensitive search.  If you want case sensitivity
    ' then change vbTextCompare to vbBinaryCompare
    If InStr(1, objItem.Name, strQuery, vbTextCompare) <> 0 Then
        %>
        <tr bgcolor="#CCFFCC">
            <td align="left" ><a href="<%= strPath & objItem.Name %>"><%= objItem.Name %></a></td>
            <td align="right"><%= objItem.Size %></td>
            <td align="left" ><%= objItem.DateCreated %></td>
        </tr>
        <%
    End If
Next 'objItem

' All done!  Kill off our object variables.
Set objItem = Nothing
Set objFolder = Nothing
Set objFSO = Nothing
%>
</table>
