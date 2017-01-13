<%
public Mibusca, Buscapor

Mibusca =trim(cstr(request.form("mibusca")))
buscapor =trim(cstr(request.form("buspor")))

public Mirut
mirut=request.querystring("Mirut")
if len(mirut)=0 then mirut =trim(cstr(request.form("Mirut")))
'if len(mirut)=0 then mirut="Nada"

':::::::::::::::::: conexion :::::::::::::::::
Dim tipodoc, mitotal, oConn, rs, sql, oConn1, rs1
Set oConn = server.createobject("ADODB.Connection")
Set oConn1 = server.createobject("ADODB.Connection")
'oConn.open "provider=Microsoft.jet.OLEDB.4.0; Data Source=" & server.mapPath("base.mdb")
oConn.open "Provider=SQLOLEDB;Data Source=serverdesa;Initial Catalog=BDFlexline;User Id=sa;Password=desakey;"
oConn1.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=todo;User Id=sa;Password=desakey;"
':::::::::::::::::: conexion :::::::::::::::::
set r=response

if len(mirut)>1 then 
	call Muestracliente()
else
	if len(Mibusca) < 1 then
		CALL nuevabus()
	else
		call listaclentes()

	end if
end if

'------------------------------------------------------------------------------------------------------
Sub listaclentes()
'on error resume next
If Buscapor="rs" then
'r.Write("razon social")
SQL="SELECT CtaCte.CodLegal, CtaCte.RazonSocial, Count(CtaCte.CodLegal) 'Contar de CodLegal' " & _
"FROM BDFlexline.flexline.CtaCte CtaCte " & _
"GROUP BY CtaCte.CodLegal, CtaCte.RazonSocial, CtaCte.Empresa, CtaCte.TipoCtaCte " & _
"HAVING (CtaCte.Empresa='DESA') AND (CtaCte.TipoCtaCte='CLIENTE') " & _
"AND (CtaCte.RazonSocial Like '%" & Mibusca & "%') AND (CtaCte.CodLegal Like '%') " & _
"ORDER BY CtaCte.CodLegal"

else
'r.Write("RUT")
SQL="SELECT CtaCte.CodLegal, max(CtaCte.RazonSocial) as RazonSocial, Count(CtaCte.CodLegal) 'Contar de CodLegal' " & _
"FROM BDFlexline.flexline.CtaCte CtaCte " & _
"GROUP BY CtaCte.CodLegal, CtaCte.Empresa, CtaCte.TipoCtaCte " & _
"HAVING (CtaCte.Empresa='DESA') AND (CtaCte.TipoCtaCte='CLIENTE') " & _
" AND (CtaCte.CodLegal Like '" & Mibusca & "%') " & _
"ORDER BY CtaCte.CodLegal"

end if

'SQL=""
Set rs=oConn.execute(Sql)

if rs.eof then r.Write("<CENTER><BR><BR>Sin Datos</CENTER>")

r.Write("<TABLE border='0'><TR>")
do until rs.eof
r.Write("<TR>")
	'r.Write("<TD>" & rs.Fields("ctacte") & " <B>" & rs.Fields("Direccion") & "</B><BR>")
'"<A HREF='documento.asp?factura=" & rsql.fields("Nfactura") &"'>"
	r.Write("<TD><FONT SIZE=1 ><A HREF='Clientes.asp?mirut=" & rs.fields("CodLegal") &"' style='text-decoration:none'>" & pc(rs.Fields("codlegal")) & "</FONT></TD></A>")
	r.Write("<TD><FONT SIZE=1 ><A HREF='Clientes.asp?mirut=" & rs.fields("CodLegal") &"' style='text-decoration:none'>" & pc(replace(rs.Fields("razonsocial")," ","&nbsp;") ) & "</FONT></a></TD>")
r.Write("</TR>")
rs.movenext
loop
r.Write("</TABLE>")
End Sub 'listaclentes()
'------------------------------------------------------------------------------------------------------
Sub Muestracliente()
on error resume next
'MiBusca=trim(cstr(request.form("mibusca")))
r.write("<basefont size='2' color='#000033' face='verdana'>")
SQL="SELECT CtaCte.CtaCte, CtaCte.RazonSocial, CtaCteDirecciones.Direccion, CtaCte.Comuna, CtaCte.Sigla, CtaCte.Ejecutivo, CtaCte.Telefono, CtaCte.CondPago, CtaCte.LimiteCredito " & _
", CtaCte.vigencia, ctacte.analisisctacte1, ctacte.analisisctacte4 " & _
"FROM BDFlexline.flexline.CtaCte CtaCte, BDFlexline.flexline.CtaCteDirecciones CtaCteDirecciones " & _
"WHERE CtaCteDirecciones.CtaCte = CtaCte.CtaCte AND CtaCteDirecciones.Empresa = CtaCte.Empresa AND CtaCteDirecciones.TipoCtaCte = CtaCte.TipoCtaCte AND ((CtaCte.Empresa='desa') AND (CtaCte.TipoCtaCte='CLIENTE') AND (CtaCteDirecciones.Principal<>'s') AND (CtaCte.CtaCte Like '" & mirut & "%')) " & _
"ORDER BY len(CtaCte.CtaCte), CtaCte.CtaCte"

Set rs=oConn.execute(Sql)
if rs.eof then
r.Write("No hay datos")
else
r.Write("<TABLE><TR>")
r.Write("<TD><FONT SIZE='2' face='verdana'>Vigencia</FONT></TD>")
vigencia=trim(rs.Fields("vigencia"))
if vigencia="B" then vigencia="Bloqueado"
if vigencia="V" then vigencia="Vigente"
r.Write("<TD colspan='2'>" & pc( vigencia ) & "</TD>")

r.Write("</TR><TR>")
r.Write("<TD><FONT SIZE='2' face='verdana'>Rut</FONT></TD>")
r.Write("<TD colspan='2'>" & pc(rs.Fields("ctacte")) & "</TD>")

r.Write("</TR><TR>")
r.Write("<TD><FONT SIZE='2' face='verdana'>Nombre</FONT></TD>")
r.Write("<TD colspan='2'>" & pc(rs.Fields("RazonSocial")) & "</TD>")

'r.Write("</TR><TR>")
'r.Write("<TD>Direccion</TD>")
'r.Write("<TD>" & rs.Fields("Direccion") & "</TD>")

r.Write("</TR><TR>")
r.Write("<TD><FONT SIZE='2' face='verdana'>Comuna</FONT></TD>")
r.Write("<TD colspan='2'>" & pc(rs.Fields("Comuna")) & "</TD>")

r.Write("</TR><TR>")
r.Write("<TD><FONT SIZE='2' face='verdana'>Sigla</FONT></TD>")
r.Write("<TD colspan='2'>" & pc(rs.Fields("Sigla")) & "</TD>")

r.Write("</TR><TR>")
r.Write("<TD><FONT SIZE='2' face='verdana'>Ejecutivo</FONT></TD>")
r.Write("<TD colspan='2'>" & pc(rs.Fields("Ejecutivo")) & "</TD>")

r.Write("</TR><TR>")
r.Write("<TD><FONT SIZE='2' face='verdana'>Telefono</FONT></TD>")
r.Write("<TD colspan='2'>" & pc(rs.Fields("telefono")) & "</TD>")

r.Write("</TR><TR>")
r.Write("<TD><FONT SIZE='2' face='verdana'>CondPago</FONT></TD>")
r.Write("<TD colspan='2'>" & pc(rs.Fields("CondPago")) & "</TD>")

r.Write("</TR><TR>")
r.Write("<TD><FONT SIZE='2' face='verdana'>LimiteCredito</FONT></TD>")
r.Write("<TD colspan='2'>" & pc(rs.Fields("LimiteCredito")) & "</TD>")

r.Write("</TR><TR>")
r.Write("<TD><FONT SIZE='2' face='verdana'>M Canal</FONT></TD>")
r.Write("<TD colspan='2'>" & pc(rs.Fields("analisisctacte1")) & "</TD>")

r.Write("</TR><TR>")
r.Write("<TD><FONT SIZE='2' face='verdana'>Descuento Maximo</FONT></TD>")
r.Write("<TD colspan='2'>" & pc(rs.Fields("analisisctacte4") & "%") & "</TD>")


'r.Write("<FONT SIZE=1 COLOR=>" )

r.Write("</TR><TR bgcolor='#E5E5E5'>")
r.Write("<TD align='center'><B><FONT SIZE='2' face='verdana'>ID_Cliente</FONT></B></TD>")
r.Write("<TD align='center'><B><FONT SIZE='2' face='verdana'>Sigla</FONT></B></TD>")
r.Write("<TD align='center'><B><FONT SIZE='2' face='verdana'>Direccion</FONT></B></TD>")
r.Write("<TD align='center'><B><FONT SIZE='2' face='verdana'>Ejecutivo</FONT></B></TD>")
do until rs.eof
r.Write("</TR><TR >")
	r.Write("<TD><A HREF='Clientes.asp?mirut=" & rs.Fields("ctacte") & "' style='text-decoration:none'>" & pc(rs.Fields("ctacte")) & "</TD></A>")
	r.Write("<TD><FONT SIZE=1  face='verdana'>" & replace(rs.Fields("sigla"    )," ","&nbsp;") & "</FONT></TD>")
	r.Write("<TD><FONT SIZE=1  face='verdana'>" & replace(rs.Fields("Direccion")," ","&nbsp;") & "</FONT></TD>")
	r.Write("<TD><FONT SIZE=1  face='verdana'>" & replace(rs.Fields("ejecutivo")," ","&nbsp;") & "</FONT></TD>")
rs.movenext
loop
r.Write("</TR></TABLE>")
'r.Write("</FONT>" )
end if

sql="SELECT RutFlex, CodSem, coddia, Dia, '' tipo, '' Terr, '' Zona, '' filen, vendedor " & _
"FROM Todo.Flexline.Rutas Rutas " & _
"WHERE (Rutas.RutFlex='" & mirut & "') " & _
"ORDER BY Rutas.CodSem, Rutas.coddia"
Set rs1=oConn1.execute(Sql)
'r.write(sql)
if rs1.eof then
	r.Write("<HR><B>Ruta Vendedor</B><BR>Sin Datos Ruta")
else
r.Write("<HR><B>Ruta Vendedor</B><BR>")
r.Write("Tipo cliente (<B>" & rs1.Fields("tipo") & "</B>)&nbsp;Territorio:<B>" & rs1.Fields("terr") & "</B>&nbsp;Zona:<B>" & rs1.Fields("Zona") & "</B>" )
r.Write("<TABLE>")
	r.Write("<TR>")
	r.Write("<TD>Semana</TD>")
	r.Write("<TD>Dia Visita</TD>")
	r.Write("<TD>Vendedor</TD>")
	r.Write("</TR>")
do until rs1.eof
	r.Write("<TR>")
	r.Write("<TD><CENTER>" & rs1.Fields("CodSem") & "</CENTER></TD>")
	r.Write("<TD><CENTER>" & rs1.Fields("dia") & "</CENTER></TD>")
	r.Write("<TD>" & rs1.Fields("Vendedor") & "</TD>")
	r.Write("</TR>")
rs1.movenext
loop
end if

end sub 'Muestracliente()
'------------------------------------------------------------------------------------------------------
function pc(texto)
	texto="<FONT SIZE='2' face='verdana'><B>" & replace(texto," ","&nbsp;") & "</B></FONT>"
	pc=texto
End function 'pc(texto)
'------------------------------------------------------------------------------------------------------
private sub nuevabus()
%>
<CENTER>
<FONT SIZE="1" COLOR="#919191">
<U>Distribucion y Excelencia S. A.</U><BR>
</FONT>
<center>
<HR>
<FONT SIZE="1" face="arial" COLOR="#000099">Busca Clientes por Nombre o Rut</FONT>
<FORM METHOD=POST ACTION="Clientes.asp">
<TABLE>
<TR>
	<TD><INPUT TYPE="text" NAME="mibusca" style="background-color: #E4EDFA"></TD>
</TR>
<TR>
	<TD>
	<INPUT TYPE="radio" NAME="buspor" id="busca1" value="rs">
	<FONT SIZE="2" face="arial" COLOR="#000066"><label for="busca1">Razon Social</label></FONT>
	</TD>
</TR>
<TR>
	<TD>
	<INPUT TYPE="radio" NAME="buspor" id="busca2" value="rut" CHECKED>
	<FONT SIZE="2" face="arial" COLOR="#000066"><label for="busca2">Codigo legal</label></FONT>
	</TD>
</TR>
<TR>
	<TD><INPUT TYPE="submit" value="Buscar"></TD>
</TR>

</TABLE>
</FORM>
<HR>
<form method='post' action='menu.asp'>
<input type=submit value="<< Menu">
</form>
<%
end sub 'nuevabus()
'------------------------------------------------------------------------------------------------------

'rs.close
oConn.close
Set rs=Nothing
Set oConn = Nothing

%>
