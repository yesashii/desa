<HEAD>
<TITLE>Incidencias</TITLE>
<style type="text/css">
	body{font-family: verdana;}
	td{font-size: 14px;}
	th{font-size: 14px;text-align: left;}
</style>
</HEAD>
<body>
<% 
':::::::::::::::::: conexion :::::::::::::::::
Dim tipodoc, mitotal, oConn, rs, sql, Minumero
Set oConn = server.createobject("ADODB.Connection")
'=========================
oConn.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=master2;User Id=sa;Password=desakey;"
'==========================
'oConn.open "provider=Microsoft.jet.OLEDB.4.0; Data Source=" & server.mapPath("base.mdb")
':::::::::::::::::: conexion :::::::::::::::::
MiNumero=trim(cstr(request.form("numero")))
if len(Minumero)=0 then MiNumero=request.querystring("Numero")
MiNumero=right("0000000000" & Minumero,10)
set r=response

':Datos de la factura
SQL="SELECT c.CtaCte, c.RazonSocial, replace(d.Direccion,',','') as direccion, c.Sigla, c.Ejecutivo, f.Vigencia " & _
"FROM handheld.flexline.CtaCte as C INNER JOIN " & _
  "handheld.flexline.CtaCteDirecciones as d ON c.CtaCte = d.CtaCte AND " & _
  "c.Empresa = d.Empresa AND c.TipoCtaCte = d.TipoCtaCte INNER JOIN " & _
  "Flexline2.dbo.Documento as f ON c.Empresa = f.Empresa AND " & _
  "d.Empresa = f.Empresa AND c.TipoCtaCte = f.TipoCtaCte AND " & _
  "c.CtaCte = f.Cliente INNER JOIN Flexline2.dbo.TipoDocumento as t ON c.Empresa = t.Empresa AND " & _
  "d.Empresa = t.Empresa AND " & _
  "f.Empresa = t.Empresa AND f.TipoDocto = t.TipoDocto AND c.TipoCtaCte = t.TipoCtaCte " & _
"WHERE (d.Principal <> 's') AND (t.FactorMonto <> 0) AND (t.Clase = 'Factura (v)') AND " & _
"(f.Numero = '" & MiNumero & "') AND (c.Empresa = 'DESA') AND (c.TipoCtaCte = 'CLIENTE')" 
Set rs=oConn.execute(Sql)
if rs.eof then
	r.write("<BR><BR><B><FONT SIZE='2' COLOR='#000066' face='Arial'>sin datos</FONT></B>")
	on error resume next
end if

r.write("<CENTER>")
r.write("<FONT SIZE='1' COLOR='#808080'>Informe Despacho</FONT><BR>")
r.write("<TABLE>")
r.write("<TR>")
	r.write("<TH>Documento</TH>")
	r.write("<TD>" & Minumero & "</TD>")
r.write("</TR>")
r.write("<TR>")
	r.write("<TH>Rut</TH>")
	r.write("<TD>" & rs.fields("CtaCte") & "</TD>")
r.write("</TR>")
r.write("<TR>")
	r.write("<TH>Nombre</TH>")
	r.write("<TD>" & rs.fields("RazonSocial") & "</TD>")
r.write("</TR>")
r.write("<TR>")
	r.write("<TH>Direc</TH>")
	r.write("<TD>" & rs.fields("Direccion") & "</TD>")
r.write("</TR>")
r.write("<TR>")
	r.write("<TH>Sigla</TH>")
	r.write("<TD>" & rs.fields("Sigla") & "</TD>")
r.write("</TR>")
r.write("</TABLE>")

r.write("<HR>")
EstadoFlex=rs.fields("Vigencia")

sql="SELECT     e.fechaentrega as fecha, e.sw_entregado as estado, e.sw_estado, e.idnrodocumento as factura " & _
"FROM         sqlserver.desaerp.dbo.TRP_INCIDENCIASENC as e " & _
"WHERE     (idnrodocumento = " & cdbl(MiNumero) & ") " & _
"ORDER BY fechaentrega"
Set rs=oConn.execute(Sql)

if rs.eof then
	r.write("Sin Informacion Incidencias<HR>")
	sql="select * from master2.flexline.fx_roadmaster where pedido=" & cdbl(MiNumero) & " "
	Set rs=oConn.execute(sql)
	if rs.eof then r.write("Sin Informacion de Ruta<HR>")
	%><TABLE bgcolor='#E8E8E8' border='1' cellpadding='2' cellspacing='0' style='border-collapse: collapse; border-width: 2;'>
	<TR>
		<TD bgcolor='#C0C0C0' background='imagenes/fd.PNG' valign='center'><B>Fecha</B></TD>
		<TD bgcolor='#C0C0C0' background='imagenes/fd.PNG' valign='center'><B>Vehiculo</B></TD>
		<TD bgcolor='#C0C0C0' background='imagenes/fd.PNG' valign='center'><B>Estado</B></TD>
	</TR><%
	Do until rs.eof
		estado ="Por Entregar"
		fecha =rs("fecha" )
		camion=rs("vehiculo")
		estado=rs("estado")
		if isnull(estado) then estado ="Por Entregar"
		if estado="1" then estado ="Entregado"

	%><TR>
		<TD bgcolor='#C0C0C0' valign='center'><B><%=fecha %></B></TD>
		<TD bgcolor='#C0C0C0' valign='center'><B><%=camion%></B></TD>
		<TD bgcolor='#C0C0C0' valign='center'><B><%=estado%></B></TD>
	</TR><%
	rs.movenext
	loop
	%></table><%
else
fecha=rs.fields("Fecha")
fecha=textfecha(fecha,"DD/MM/YYYY")
r.write("<TABLE bgcolor='#E8E8E8' border='1' cellpadding='2' cellspacing='0' style='border-collapse: collapse; border-width: 2;'>")
r.write("<TR>")
	r.write("<TD bgcolor='#C0C0C0' background='imagenes/fd.PNG' valign='center'><B>Fecha</B></TD>")
	r.write("<TD bgcolor='#C0C0C0' background='imagenes/fd.PNG' valign='center'><B>Vehiculo</B></TD>")
	r.write("<TD bgcolor='#C0C0C0' background='imagenes/fd.PNG' valign='center'><B>SW&nbsp;Estado</B></TD>")
	r.write("<TD bgcolor='#C0C0C0' background='imagenes/fd.PNG' valign='center'><B>Incidencia</B></TD>")
	r.write("<TD bgcolor='#C0C0C0' background='imagenes/fd.PNG' valign='center'><B>Comentario</B></TD>")
r.write("</TR>")
on error resume next
do until rs.eof
	fecha=rs.fields("Fecha")
	fecha=textfecha(fecha,"DD/MM/YYYY")
	ippdesc="0"
	OBS="&nbsp;"
	estado="Por Entregar"
	if trim(rs.fields("estado")) = "N" then estado = "Por Entregar"
	if trim(rs.fields("estado")) = "S" then estado = "Entregado OK"
	if estado="Por Entregar" then
		if cdate(date()) > cdate(fecha) then estado = "No Entregado"
	end if

	if rs("sw_estado")="E" then sw_estado="Entregado"
	if rs("sw_estado")="A" then sw_estado="Anulado"
	if rs("sw_estado")="N" then sw_estado="No Entregado"
	if rs("sw_estado")="R" then sw_estado="Redespacho"

	sql="SELECT * " & _
	"FROM sqlserver.desaerp.dbo.TRP_INCIDENCIASDET " & _
	"where idempresa=1 and idnrodocumento=" & cdbl(MiNumero) & " and fechaentrega= " & rs.fields("Fecha")
	'r.write(sql)
	set rs2=oConn.execute(Sql)
	if not rs2.eof then
		OBS=rs2.fields("comentario")
		ippdesc=rs2.fields("idincidencia")
	end if

	if len(obs)=0 then OBS="&nbsp;"

	if isnull(ippdesc) then ippdesc=""
	if len(ippdesc)>0 then
		sqlipp="select i.nombre from sqlserver.desaerp.dbo.DIM_INCIDENCIAS as i where i.idincidencia=@ipp"
		set rs1=oConn.execute(replace(sqlipp,"@ipp",ippdesc))
		if not rs1.eof then 
			ippdesc=rs1.fields(0)
		else
			ippdesc=""
		end if
	end if
	if ippdesc="0" then ippdesc=""

	'rmfecha=textfecha(cstr(rs("fecha")),"DD/MM/YYYY")' 00:00:00")
	fechax=rs.fields("Fecha")
	rmfecha=textfecha(fechax,"YYYY-MM-DD 00:00:00")

	sql="SELECT Pedido, Vehiculo, Fecha " & _
	"FROM Flexline.FX_ROADMASTER " & _
	"WHERE (Pedido = N'" & cdbl(MiNumero) & "') AND (Fecha = CONVERT(DATETIME, '" & rmfecha & "', 102))"
	'r.write("<HR>" & fechax & ":" & rmfecha & ":" & sql)
	set rs3=oConn.execute(sql)
	if not rs3.eof then vehiculo=rs3("vehiculo")

	r.write("<TR>")
		r.write("<TD><CENTER>" & fecha & "</CENTER></TD>")
		r.write("<TD><CENTER>" & vehiculo & "</CENTER></TD>")
		'r.write("<TD><CENTER>" & estado & "</CENTER></TD>")
		r.write("<TD><CENTER>" & sw_estado & "</CENTER></TD>")
		r.write("<TD><CENTER>" & pc(ippdesc,0) & "</CENTER></TD>")
		r.write("<TD>" & pc(OBS,0) & "</TD>")
	r.write("</TR>")

rs.movenext
loop

r.write("</TABLE>")

'EstadoFlex
if EstadoFlex="S" then EstadoFlex ="Factura Activa"
if EstadoFlex="N" then EstadoFlex ="*Nota de credito Asociada"
if EstadoFlex="A" then EstadoFlex ="Factura NULA"
r.write(EstadoFlex)
end if

r.write("<BR><INPUT TYPE='button' value='Volver' onclick='history.back()'>")
r.write("</CENTER></body>")

rs.close
oConn.close
Set rs=Nothing
Set oConn = Nothing

'------------------------------------------------------------------------------------------------
Function PC(sString, mzise) 
  Dim sWhiteSpace, bCap, iCharPos, sChar 
  sWhiteSpace = Chr(32) & Chr(9) & Chr(13) 
  sString = LCase(sString)
  bCap = True 
  For iCharPos = 1 to Len(sString) 
    sChar = Mid(sString, iCharPos, 1) 
    If bCap = True Then 
      sChar = UCase(sChar) 
    End If 
    ProperCase = ProperCase + sChar 
    If InStr(sWhiteSpace, sChar) Then 
      bCap = True 
    Else 
      bCap = False 
    End If 
  Next  
PC="<FONT SIZE='" & Mzise & "' face='arial' COLOR='#000033'>" & replace(ProperCase," ","&nbsp;") & "</FONT>"
PC=replace(replace(PC,"%20"," "),"***",chr(34))
if mzise=0 then PC=ProperCase
End Function 

'------------------------------------------------------------------------------------------------
Function textfecha(lafecha,formato)
	formato=Ucase(formato)
	lafecha=cstr(trim (lafecha))
	if len(lafecha)<>8 then
		textfecha="Err:largo:" & lafecha
		exit Function
	end if
	if not isnumeric(cdbl(lafecha)) then
		textfecha="Err:Numer:" & lafecha
		exit Function
	end if
		nYYYY=left(lafecha,4)
		nMM  =right(left(lafecha,6),2)
		nDD  =right(lafecha,2)
		formato=replace(formato,"YYYY",nYYYY)
		formato=replace(formato,"MM"  ,nMM  )
		formato=replace(formato,"DD"  ,nDD  )
		textfecha = cstr(formato)
End Function
'------------------------------------------------------------------------------------------------
%>