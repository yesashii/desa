<%
on error resume next
vend 	= trim(request.querystring("vend"))
tipPed 	= trim(request.QueryString("tipPed"))
cliente = trim(request.querystring("cliente"))
vend=replace(vend,"?","Ñ")'Error blackberry
'vend=replace(vend,"NN","Ñ")'Error blackberry
id_porfolio=trim(request.QueryString("id_porfolio"))
empresa=trim(request.QueryString("empresa"))
sw_iva		= request.QueryString("sw_iva")
tipodocto   = request.QueryString("tipodocto")
'response.write("<BR>empresa : " & empresa)
response.write("<BR>ID porfolio : " & id_porfolio)
':::::::::::::::::: conexion :::::::::::::::::
Private tipodoc, mitotal, oConn, rs, sql, rs1
Set oConn = server.createobject("ADODB.Connection")
oConn.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=handheld;User Id=sa;Password=desakey;"

SQL="SELECT CtaCte.CodLegal, CtaCte.RazonSocial, CtaCteDirecciones.Direccion, CtaCte.CtaCte, CtaCte.Giro, CtaCte.CondPago, CtaCte.Comuna, CtaCte.Telefono, CtaCte.LimiteCredito, CtaCte.PorcDr1, ctacte.vigencia, CtaCte.porfolio1, CtaCte.porfolio2 " & _
"FROM handheld.flexline.CtaCte CtaCte, handheld.flexline.CtaCteDirecciones CtaCteDirecciones " & _
"WHERE CtaCteDirecciones.CtaCte = CtaCte.CtaCte AND CtaCteDirecciones.Empresa = CtaCte.Empresa AND CtaCteDirecciones.TipoCtaCte = CtaCte.TipoCtaCte AND ((CtaCte.Empresa='DESA') AND (CtaCte.TipoCtaCte='CLIENTE') AND (CtaCte.Ejecutivo='" & vend & "') AND (CtaCte.CodLegal='" & cliente & "') AND (CtaCteDirecciones.Principal<>'s'))"

if lcase(id_porfolio)="porfolio2" then
	'response.write("<B>PORFOLIO2</B>")
	if len(empresa)>0 then 
		SQL=replace(Ucase(SQL),"EMPRESA='DESA'","EMPRESA='" & empresa & "'")
	else
		SQL=replace(Ucase(SQL),"EJECUTIVO","PORFOLIO2")
	end if

	'SQL2=replace(ucase(SQL2),"VENDEDOR" ,"PORFOLIO2")
else
	'SQL=replace(ucase(SQL),"EJECUTIVO","PORFOLIO1")
	'response.write("else")
end if

if ucase(empresa)="DESAZOFRI" then 
		SQL=replace(Ucase(SQL),"EMPRESA='DESA'","EMPRESA='" & empresa & "'")
end if

''77
'if vend="CARLOS VALDEBENITO" then
'SQL="SELECT CtaCte.CodLegal, CtaCte.RazonSocial, CtaCteDirecciones.Direccion, CtaCte.CtaCte, CtaCte.Giro, CtaCte.CondPago, CtaCte.Comuna, " & _
'"CtaCte.Telefono, CtaCte.LimiteCredito, CtaCte.PorcDr1, LEN(CtaCte.CtaCte) AS Expr1, ctacte.vigencia " & _
'"FROM flexline.CtaCte CtaCte INNER JOIN " & _
'"flexline.CtaCteDirecciones CtaCteDirecciones ON " & _
'"CtaCte.CtaCte = CtaCteDirecciones.CtaCte AND CtaCte.Empresa = CtaCteDirecciones.Empresa AND " & _
'"CtaCte.TipoCtaCte = CtaCteDirecciones.TipoCtaCte " & _
'"WHERE (CtaCte.Empresa = 'DESA') AND (CtaCte.TipoCtaCte = 'CLIENTE') AND (CtaCte.CodLegal='" & cliente & "') " & _
'"AND (CtaCteDirecciones.Principal <> 's') " & _
'"ORDER BY LEN(CtaCte.CtaCte), CtaCte.CtaCte"
'end if


Set rs=oConn.execute(Sql)
'response.write(sql)
%>
<html>

<head>
<meta http-equiv="Content-Language" content="es-cl">
<title>Seleccione Local</title>
</head>

<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0">
<%
if rs.eof then 
'response.write("<BR><BR>" & sql & "<BR><BR>")
response.write("<p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p><p align='center'>Error en Maestro de Clientes</p><p align='center'>No hay direcci&oacute;n registrada o Cliente Mal Creado</p>")
response.write("<p>&nbsp;</p><p align='center'><input type='button' value='&nbsp;&nbsp;&nbsp;Volver&nbsp;&nbsp;&nbsp;' onclick='history.back()'>&nbsp;</p>")
else
%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; border-width: 0" bordercolor="#111111" width="100%" height="141">
  <tr>
    <td width="100%" style="border-style: none; border-width: medium" align="center" height="29">
    <p align="center">
    <b><font face="Arial" size="2" color="#333333">Sistema Preventa</font></b></td>
  </tr>
    <tr>
    <td width="100%" style="border-style: none; border-width: medium" align="center" height="13">
    <font size="2" face="Arial">&nbsp;</font></td>
  </tr>
  <tr>
    <td width="100%" style="border-style: none; border-width: medium" align="left" height="13">
    <b><font size="2" face="Arial"><% = replace(rs.fields("razonsocial")," ","&nbsp;") %></font></b></td>
  </tr>
  <tr>
    <td width="100%" style="border-style: none; border-width: medium" align="left" height="13">
    <font size="2" face="Arial"><% = replace("Rut : " & rs.fields("codlegal")," ","&nbsp;") %></font></td>
  </tr>
  <tr>
    <td width="100%" style="border-style: none; border-width: medium" align="center" height="93">
    <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
      <tr>
        <td width="100%" bgcolor="#000080">
        <font face="Arial" size="2" color="#FFFFFF">Locales</font></td>
      </tr>
<%
do until rs.eof
	if ucase(trim(rs.fields("porfolio1")))=ucase(vend) then
		'if len(trim(rs.fields("porfolio2")))>0 then id_porfolio="porfolio1"
		if rs.fields("porfolio2")="NO ASIGNADO" then id_porfolio=""
	end if

	response.write("<tr>")
	response.write("<td width='100%'><font face='Arial' size='2'>")
	if tipped = "S" then
		if ucase(trim(rs.fields("vigencia")))<>"S" then
			response.write("Cliente NO VIGENTE")
		else
	response.write("<a href='verifica_cliente.asp?vend=" & vend & "&cliente=" & rs.fields("ctacte") & "&id_porfolio=" & id_porfolio & "&empresa=" & empresa & "&tipodocto=" & tipodocto & "&sw_iva=" & sw_iva & "' style='text-decoration: none'>(" & RIGHT(rs.fields("ctacte"),2) & ")&nbsp;" & replace(rs.fields("direccion"),"","&nbsp;") & "</a>")
		end if
	end if
	if tipped = "N" then
	response.write("<a href='noventa/noventa.asp?vend=" & vend & "&cliente=" & rs.fields("ctacte") & "&id_porfolio=" & id_porfolio & "&empresa=" & empresa & "&tipodocto=" & tipodocto & "&sw_iva=" & sw_iva & "' style='text-decoration: none'>(" & RIGHT(rs.fields("ctacte"),2) & ")&nbsp;" & replace(rs.fields("direccion"),"","&nbsp;") & "</a>")
	
	end if
response.Write("</font></td>")
response.write("</tr>")
rs.movenext
loop
'response.Write(vend)
end if
%>
	</table>
    </td>
  </tr>
  </table>
<p align="center"><font face="Arial" size="1" color="#808080"></font></p>
</body>
</html>