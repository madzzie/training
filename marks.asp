<%@ LANGUAGE=VBSCRIPT %>
<%Dim Db, Rs, i

set db = server.createobject("adodb.connection")
set rs = server.createobject("adodb.recordset")
db.open "Provider=MSDAORA.1;data source=XE;user id=gbk;password=intech"
rs.open "select * from kutty",db

<html><head><title>:: Statewise Branches ::</title></head>
<body>

<p align=center style='font-size:14pt;color:purple;margin:0'>Marks Sheet</p>

<table width=50% align=center border=1 bordercolor=#e0e0e0 style='border-collapse:collapse;color:blue;font-size:9pt;'>
<tr bgcolor=khaki style='font-weight:bold;color:black'>
<td>Sl No</td>
<td>Suject</td>
<td>Marks</td>
<td>Grade</td>
</tr>

<%while not rs.eof%>
	<tr>
		<td><%=rs("slno")%></td>
		<td><%=rs("subject")%></td>
		<td><%=rs("marks")%></td>
		<td><%=rs("grade")%></td>
	</tr>
	<%rs.movenext%>
<%wend%>

<tr bgcolor=khaki style='font-weight:bold;color:black'><td>Total Marks</td><td align=center colspan=3>&nbsp;</td></tr>
</table>
</body>
</html>