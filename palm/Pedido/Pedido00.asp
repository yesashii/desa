<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<%
vend 		= trim(request.querystring("vend"))
tipPed 		= trim(request.QueryString("tipPed"))
buscadnom 	= trim(request.form("buscar"))
vend        = replace(vend,"?","Ñ")'Error blackberry
'vend        = replace(vend,"NN","Ñ")'Error blackberry
id_porfolio = trim(request.QueryString("id_porfolio"))
nuser       = trim(request.querystring("nuser"))
empresa     = trim(request.querystring("empresa"))
sw_iva		= request.form("ck_iva")
tipodocto   = request.form("idtipodocto")
'if len(tipodocto)=0 then tipodocto=trim(request.querystring("idtipodocto"))
'response.write(nuser)
'response.write("<BR>sw_iva : " & sw_iva)
'response.write("<BR>tipodocto : " & tipodocto)
'response.write(id_porfolio)
'if len(id_porfolio)<1 then id_porfolio="Porfolio1"
response.write("<BR>vend : " & vend)
response.write("<BR>empre : " & empresa)
'response.write("<A HREF="">hola</A>")
'response.write("<BR>tipPed : " & tipPed)

on error resume next
':::::::::::::::::: conexion :::::::::::::::::
Private tipodoc, mitotal, oConn, rs, sql, rs1
Set oConn = server.createobject("ADODB.Connection")
oConn.ConnectionTimeOut = 0
oConn.CommandTimeout = 0
'oConn.open "provider=Microsoft.jet.OLEDB.4.0; Data Source=" & server.mapPath("base.mdb")
'oConn.open "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=todo;User Id=sa;Password=desakey;"
oConn.open "Provider=SQLOLEDB;Data Source=sqlserver;Initial Catalog=todo;User Id=sa;Password=desakey;"

':::Busca por cliente
'anterior localhost
SQL1="SELECT TOP 100 CodLegal, RazonSocial, COUNT(CodLegal) 'Contar de CodLegal'" & _
"FROM handheld.flexline.CtaCte C " & _
"WHERE (Empresa = 'DESA') AND (TipoCtaCte = 'CLIENTE') AND (Ejecutivo = '" & vend & "') AND (vigencia='S') AND" & _
"(CodLegal + RazonSocial + sigla LIKE '%" & buscadnom & "%')" & _
"GROUP BY CodLegal, RazonSocial" '& _
'"ORDER BY RazonSocial"
'response.write(SQL1)

'::: Busca por ruta
SQL2="SELECT *, rutflex as 'CodLegal', getdate() as fecha FROM handheld.Flexline.PDA_RUTA_HOY2 " & _
"WHERE (Vendedor = '" & vend & "')"
'SQL2="exec flexline.consulta_ruta @ejecutivo='" & vend  & "'"
'response.write(sql2)

':::: supercadoA
SQL3="SELECT CtaCte.CodLegal, CtaCte.RazonSocial " & _
"from  handheld.Flexline.ctacte as ctacte " & _
"where (ctacte.ejecutivo = 'CARLOS VALDEBENITO') " & _
"GROUP BY CtaCte.CodLegal, CtaCte.RazonSocial"
'Detalle del listado


'on error resume next '*******************************************

if ucase(empresa)="DESAZOFRI" then
	SQL1=replace(SQL1,"'DESA'","'DESA'")
	SQL2=replace(SQL2,"'DESA'","'DESA'")
	SQL3=replace(SQL3,"'DESA'","'DESA'")
end if

if lcase(id_porfolio)="porfolio2" then
	response.write("<B>PORFOLIO2</B>")
	
	'SQL1=replace(Ucase(SQL1),"EJECUTIVO","PORFOLIO2")
	SQL1=replace(Ucase(SQL1),"'DESA'","'LACAV'")
	'SQL2=replace(ucase(SQL2),"VENDEDOR" ,"PORFOLIO2")*********
'else
	'SQL1=replace(ucase(SQL1),"EJECUTIVO","PORFOLIO1")
'	response.write("<BR><BR>" & SQL1 & "<BR><BR>" & SQL2 & "<BR><BR>")
end if

response.flush()

'Estado pedido Dia
nfecha=right(year(date),2) & right("00" & month(date),2) & right("00" & day(date),2)
SQLX="SELECT vendedor, estado, fecha, nota, Cliente " & _
"FROM handheld.Flexline.FX_PEDIDO_PDA " & _
"WHERE (fecha = N'" & nfecha & "') AND (vendedor = N'" & vend & "') " & _
"order by Cliente, estado "
Set rs1=oConn.execute(SQLX)

response.flush()

if len(buscadnom)=0 then
	Set rs=oConn.execute(SQL2)
	if not rs.eof then
		Midia="&nbsp;"
		if rs.fields("coddia")=1 then Midia="Lun"
		if rs.fields("coddia")=2 then Midia="Mar"
		if rs.fields("coddia")=3 then Midia="Mie"
		if rs.fields("coddia")=4 then Midia="Jue"
		if rs.fields("coddia")=5 then Midia="Vie"
		if rs.fields("coddia")=6 then Midia="Sab"
		if rs.fields("coddia")=7 then Midia="Dom"
		DetList= Midia & "-Semana:" & rs.fields("codsem")
		'response.write(SQL2)
	else
		Set rs=oConn.execute(SQL1)
		DetList= "todos"
	end if
Else
	Set rs=oConn.execute(SQL1)
	DetList= "busca:" & buscadnom
	'response.write(SQL1)
end if 


'77
'if ucase(trim(vend))=ucase(trim("CARLOS VALDEBENITO")) then
'	Set rs=oConn.execute(SQL3)
'	DetList= "todos"
'end if

%>
<head>
<title>Seleccionar Cliente</title>
</head>

<body topmargin="0" leftmargin="0">

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; border-width: 0" bordercolor="#111111" width="100%">
  <tr>
    <td style="border-style: none; border-width: medium">
		<CENTER><B>
		<!-- <img src="../logo.JPG" width="32" height="20"> -->
			<font face="Arial" size="2" color="#333333">
				Sistema Preventa
			</font>
		</B></CENTER>
     </td>
  </tr>
  <tr>
    <td width="100%" style="border-style: none; border-width: medium" align="center">
    <table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse">
	  <tr>
	  	<td width="100%" align="right" valign="middle">
			<form method="post" action="">
				<FONT SIZE="2" COLOR="#808080">Busqueda</FONT>
				<input name="buscar" type="text" size="20">
	  			<input type="submit" name="buscanom" value="buscar">
				<INPUT TYPE="hidden" name="idtipodocto" value="<%=tipodocto%>">
				<INPUT TYPE="hidden" name="ck_iva" value="<%=sw_iva%>">
			</form>
		</td>
	</tr>
	<tr>
        <td width="100%" align="center" valign="middle" bgcolor="#666633">
          <font face="Arial" size="2" color="#FFFFFF">
			List.&nbsp;Clientes&nbsp;(<%=DetList %>)
		  </font></td>
      </tr>
<%
'instr
'if len(trim(cstr(request.form("buscanom")))) then
'SQL1="SELECT CodLegal + RazonSocial AS busca FROM flexline.CtaCteCtaCte.CodLegal, CtaCte.RazonSocial where busca like'%" & trim(cstr(request.form("buscar"))) & "%'"
'end if
If rs.bof then response.write("<BR>La busqueda no produjo ninguna Concidencia")
response.flush()
do until rs.eof
	Estadoped=""

	if instr(rs.fields("CodLegal")," ")=0 then
		Milink="Pedido01.asp"
		LocalCliente=""
	else
		Milink="verifica_cliente.asp"
		LocalCliente="(" & mid(rs.fields("CodLegal"),instr(rs.fields("CodLegal")," ")+1) & ")"
		'busca estado
		Do until rs1.eof or rs1.bof
			If trim(rs.fields("CodLegal"))=trim(rs1.fields("Cliente"))then 
				if rs1.fields("estado")="noventa" then
					Estadoped="<B>NV</B>"
				else
					Estadoped="<B>PD</B>"
				end if
				
				if rs1.fields("estado")="OK" then
					Estadoped="<B>PD</B>"
				end if
			end if
		rs1.movenext
		loop
		rs1.MoveFirst
	end if
	
	if tipPed="N" then Milink="noventa/noventa.asp"
	'if lcase(id_porfolio)="porfolio2" then Milink=replace(Milink,".asp",".asp")

	CodLegal = rs.fields("CodLegal")
	response.write("<tr>")
	response.write("<td width='100%' align='center'><FONT face='Arial' SIZE='2'>&nbsp;" & _
	"<a href='" & Milink & "?vend=" & vend & "&cliente=" & CodLegal & "&tipPed=" & tipPed & "&id_porfolio=" & id_porfolio & "&empresa=" & empresa & "&tipodocto=" & tipodocto & "&sw_iva=" & sw_iva & "' style='text-decoration: none'>" & Estadoped & LocalCliente & " " &  _
	replace(rs.fields("RazonSocial")," ","&nbsp;") & _ 
	"</a></FONT></td>")
	response.write("</tr>")
	response.flush()
rs.movenext
loop
%>
    </table>
    </td>
  </tr>
</table>
	<font face="Arial" size="1" color="#808080">
		&nbsp;NV=No Venta<BR>
		&nbsp;PD=Pedido<BR>
		&nbsp;* Mientras Menos letras pongas en el buscador, mas concidencias encuentras

	</font>
</body>
</html>