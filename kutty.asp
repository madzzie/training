<%@ LANGUAGE=VBSCRIPT %>
<%Dim Db, Rs, i
set db = server.createobject("idbi_dal.oradal")
set rs = db.getrs("SELECT STATE, Count(*) cnt FROM SOL GROUP BY STATE ORDER BY Count(*) DESC")%>

<html><head><title>:: Statewise Branches ::</title></head>
<body>

<p align=center style='font-size:14pt;color:purple;margin:0'>Statewise listing of branches</p>

<table width=50% align=center border=1 bordercolor=#e0e0e0 style='border-collapse:collapse;color:blue;font-size:9pt;'>
<tr bgcolor=khaki style='font-weight:bold;color:black'>
<td>State</td>
<td align=center>Branches</td>
</tr>

<%while not rs.eof%>
	<tr>
	<td><%=rs("state")%></td>
	<td align=center><%=rs("cnt")%></td>
	</tr>
	<%i = i + cdbl(rs("cnt")) 
	rs.movenext%>
<%wend%>

<tr bgcolor=khaki style='font-weight:bold;color:black'><td>Total Branches</td><td align=center><%=i%></td></tr>
</table>
</body>
</html>