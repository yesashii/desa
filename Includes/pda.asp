<%
'-------------------------------------------------------------------------------------------
sub encabezado(nuser, tipocab)
empresa=request.querystring("empresa")
'response.write(empresa)
nuser=cint(nuser)
nvendedor=""
sql="select nombre from SQLSERVER.Desaerp.dbo.DIM_vendedores where idvendedor=" & nuser & " and nombre <>'CLAUDIO CROQUEVIELLE'"
if len(empresa)>0 then 
	if empresa="DESAZOFRI" then idempresa=3
	if empresa="DESA" then idempresa=1
	if empresa="LACAV" then idempresa=4
	sql="select nombre from SQLSERVER.Desaerp.dbo.DIM_vendedores where idvendedor=" & nuser & " and nombre <>'CLAUDIO CROQUEVIELLE' and idempresa='" & idempresa & "'"
end if
Set rs=oConn.execute(Sql)
if not rs.eof then nvendedor=rs.fields(0)
rs.close	
%><TABLE width="100%" border='0' cellpadding='0' cellspacing='0'>
<TR bgcolor="#666633">
	<TD align="center"><B><FONT SIZE="2" face="Arial" COLOR="#FFFFFF"><%=pc(nuser,0)%></B></FONT></TD>
	<TD align="center"><B><FONT SIZE="2" face="Arial" COLOR="#FFFFFF"><%=pc(nvendedor,0)%></B></FONT></TD>
	<TD align="center"><B><FONT SIZE="2" face="Arial" COLOR="#FFFFFF"><%=pc(tipocab,0)%></B></FONT></TD>
</TR>
</TABLE><%
End sub 'encabezado(tipocab)
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
%>