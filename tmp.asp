<%@language=vbscript%>
<%DIM db,rs, clr
set db = server.createobject("adodb.connection")
set rs = server.createobject("adodb.recordset")
db.open "Provider=MSDAORA.1;data source=XE;user id=gbk;password=intech"
rs.open "SELECT bscode, nvl(base_amt,0) base_amt, nvl(sub_amt,0) sub_amt, nvl(moc,0) moc, nvl(de_amt,0) de_amt, nvl(final_amt,0) final_amt, nvl(prev_amt,0) prev_amt FROM IBANK_BS_CONS order by bscode",db%>
<html>
	<head><title>TEMP</title></head>
<body>
<h3 style='color:green;' align=center><b>IDBI Bank, Head Office, IDBI Tower, World Trade Center, Mumbai</b></h3>
<table border=1 width=100% bordercolor=#e0e0e0 style='bgcolor:lightpink;border-collapse:collapse;color:purple;font-size:10pt'>

<tr bgcolor=pink>
	<td>BS CODE</td>
	<td>Base Amt</td>
	<td>Sub_amt</td>
	<td>MOC</td>
	<td>DE_amt</td>
	<td>Final_amt</td>
	<td>Prev_amt</td>
</tr>

<%while not rs.eof%>
	<tr color=red align=right>
		<td align=center><%=rs("BSCODE")%></td>
		<%if cdbl(rs("BASE_AMT")) < 0 then clr="red" else clr="blue"%>
		<td style='color:<%=clr%>'><%=rs("BASE_AMT")%></td>
		<%if cdbl(rs("sub_AMT")) < 0 then clr="red" else clr="blue"%>
		<td style='color:<%=clr%>'><%=rs("SUB_AMT")%></td>
		<%if cdbl(rs("moc")) < 0 then clr="red" else clr="blue"%>
		<td style='color:<%=clr%>'><%=rs("MOC")%></td>
		<%if cdbl(rs("dE_AMT")) < 0 then clr="red" else clr="blue"%>
		<td style='color:<%=clr%>'><%=rs("DE_AMT")%></td>
		<%if cdbl(rs("final_AMT")) < 0 then clr="red" else clr="blue"%>
		<td style='color:<%=clr%>'><%=rs("FINAL_AMT")%></td>
		<%if cdbl(rs("prev_AMT")) < 0 then clr="red" else clr="blue"%>
		<td style='color:<%=clr%>'><%=rs("PREV_AMT")%></td>
	</tr>
	<%rs.MoveNext
wend%>

<tr bgcolor=pink><td colspan=7>&nbsp;</td></tr>
</table>
</html>