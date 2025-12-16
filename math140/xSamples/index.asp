<html>
<body>

<p>
<%
dim course
course=Request.QueryString("course")
%>
Sample Exams and Quizzes <% If course<>"" then response.write("for " & course) %>
</p>

<%
dim c
dim i
set nl=server.createobject("MSWC.Nextlink")
c = nl.GetListCount("catalog.txt")
i = 1
%>
<ul>
<%do while (i <= c) %>
<li><a href="<%=nl.GetNthURL("catalog.txt", i)%>">
<%=nl.GetNthDescription("catalog.txt", i)%></a>
<%
i = (i + 1)
loop
%>
</ul>
<p>
<a href="../">Return to Course Webpage</a>
</body>
</html>
