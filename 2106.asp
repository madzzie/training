<%@language=vbscript%>
<%Dim db,rs,i
set db=server.createobject("adodb.connection")
set rs=server.createobject("adodb.recordset")
db.open "provider=MSDAORA.1;data source=XE;user id=gbk;password=intech"
rs.open "select * from report",db
%>

<html><head><title>Report Card</title></head>
<body>

<p align=center style='font-size:15pt;font-weight:bold;color:turquoise;margin:0'>Marks Sheet</p>

<table width=50% align=center border=1 bordercolor=#e0e0e0 style='border-collapse:collapse;color:purple;font-size:10pt;'>
<tr bgcolor=pink style='font-weight:bold;color:brown'>
<td>Sl no</td>
<td>Suject</td>
<td>Marks</td>
<td>Grade</td>
</tr>

<%while not rs.eof%>
	<tr>
		<td><%=i%></td>
		<td><%=rs("Subject")%></td>
		<td><%=rs("Marks")%></td>
		<td><%=rs("Grade")%></td>
	</tr>
	<%i=i+1
	rs.movenext%>
<%wend%>

<tr bgcolor=pink style='font-weight:bold;color:brown'><td>GPA</td><td align=right colspan=3>9.8</td></tr>
</table>
</body>
</html>