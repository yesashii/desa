<html>
<head>
<title>Traspaso</title>
</head>
<body bgcolor="#FFFFFF">
<% 'on error resume next
':::::::::::::::::: conexion :::::::::::::::::
private tipodoc, mitotal, oConn, rs, sql, rs1, SQL2, oConn2
Set oConn  = server.createobject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.RecordSet")
oConn.ConnectionTimeOut = 0
oConn.CommandTimeout = 0

oConn.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=handheld;User Id=sa;Password=desakey;"

vend=request.querystring("vend")
vend=replace(vend,"?","Ñ")'Error blackberry
'vend=replace(vend,"NN","Ñ")'Error blackberry

laempresa=request.querystring("empresa")
sql="select * from flexline.fx_vende_man where nombre='" & vend & "'"
set rs1=oConn.execute(SQL)

if not rs1.eof or not rs1.bof then 
	nuser=rs1.fields("nu_2")
	empresa=rs1.fields("empresa")
end if

nempresa=1
idporfolio=request.querystring("id_porfolio")
'if len(idporfolio)<1 then idporfolio="porfolio1"

response.write("idporfolio : " & idporfolio)
response.write("<BR>len : " & len(idporfolio) )

if idporfolio="porfolio1" then empresa="DESA"
if idporfolio="porfolio2" then empresa="LACAV"
if empresa="DESA"  then nempresa=1
if empresa="LACAV" then nempresa=4

if laempresa="DESAZOFRI" then
	empresa="DESAZOFRI"
	nempresa=3
end if

if len(nuser)<1 then nuser=request.form("nuser")
if len(nuser)<1 then nuser=request.querystring("nuser")
if len(nuser)<1 then nuser=request.cookies("fx_nusuario")
idbodega=Request.Form("idbodega")

nuser=Right("000" & nuser,3)
'response.write(nuser)
if request.cookies("fx_pedidorep")="si" then 
	%>
	<p align="center"><b>El Pedido ya fue ingresado</b></p>
	<form method="post" action="pedido.asp?nuser=<%= nuser %>">
	<p align="center"><input type="submit" value="Salir"></p></form>
	<%
else
Response.cookies("fx_pedidorep")="si"
'========================================================================================
'=============      CAPTURA PEDIDOS PDA    -    version : 1.0              ==============
'========================================================================================
OC =trim(cstr(request.form("oc"  )))
if len(oc )=0 then oc="0000"
OBS=trim(cstr(request.form("obs" )))
OBS2=trim(cstr(request.form("obs2" )))
	if len(OBS2)=0 then OBS2 = " "
	if len(OBS)=0 then OBS = " "
fc_des = left(request.Form("fechaentrega"),8)
if len(trim(fc_des))=0 then
	fc_des=null
else
	hoy= year(date()) & right("00" & month(date()),2) & right("00" & day(date()),2)
	if cdbl(fc_des)=<cdbl(hoy) then fc_des=hoy
end if

'response.Write(fc_des)
dim Midetalle(11,2)
':::::: MATRIZ - MiDetalle ::::::
'Producto | Cantidad | Descuento
'--------------------------------
'    0    |    1     |    2

For Each elemento In request.querystring
	if len(elemento)=3 then
		if left(elemento,1)="p" then Micol=0
		if left(elemento,1)="c" then Micol=1
		if left(elemento,1)="d" then Micol=2
		Mifil=cdbl(right(elemento,2)-1)
		if len(request.querystring(elemento))>0 then
			Midetalle(Mifil,Micol)=request.querystring(elemento)
		else
			if left(elemento,1)="p" then
				Midetalle(Mifil,Micol)=""
			else
				Midetalle(Mifil,Micol)=0
			end if
		end if
	END IF
Next


'********************************************************
'********************************************************	
vend=request.QueryString("vend")
vend=replace(vend,"?","Ñ")'Error blackberry
if request.querystring("cliente")="81788500-4 2" then
	vend ="CALL CENTER 1"
	nuser="17"
end if
if request.querystring("cliente")="81788500-4 3" then
	vend ="DAVID CARRILLO"
	nuser="35"
end if

SQL2="select *, correlativo_notapedido as notap from handheld.dbo.ped_notafolios where idvendedor=" & cint(nuser) & " and idempresa=" & nempresa & ""
'response.write SQL2
set rs1=oConn.execute(SQL2)
	if rs1.eof and rs1.bof  then
		nota=0
	else
		nota=rs1.fields("notap")
	end if
	rs1.close
nwnota= (Cdbl(nota) + 1)
nwnota= (Cdbl(right(nota,4)) + 1)

'*********************************************************
'*********************************************************
'on error resume next
	'vend=request.querystring("vend")
	'vend=replace(vend,"?","Ñ")'Error blackberry
	sql="select top 1 * FROM handheld.flexline.FX_PEDIDO_PDA"
'	Set rs=oConn.execute(Sql)
'	rs.Close
	rs.Open SQL, oConn, 1, 3
	rs.AddNew
	rs.Fields("vendedor"             ) = vend'replace(vend,"NN","Ñ")
	Mifecha=right(year(date),2) & right("00" & month(date),2) & right("00" & day(date),2)
	Mifechae=right(year(date+1),4) & right("00" & month(date+1),2) & right("00" & day(date+1),2)
	mihora=right("00" & hour(time),2) & right("00" & minute(time),2) & right("00" & second(time),2)
	sw_iva   =request.querystring("sw_iva"   )
	tipodocto=request.querystring("tipodocto")
	if tipodocto="001" then sw_iva="1"
	if len(sw_iva)=0 then sw_iva="0"
	if ucase(sw_iva)=ucase("on") then sw_iva="1"

	if isnull(fc_des) then fc_des=Mifechae
	if len(trim(fc_des))=0 then fc_des=Mifechae
	'if empresa="LACAV" then empresa="DESA"
	rs.Fields("fecha"                ) = mifecha
	rs.Fields("hora"                 ) = mihora
	rs.Fields("cliente"              ) = request.querystring("cliente")
	rs.Fields("nota"                 ) = nwnota
	'rs.Fields("oc"                   ) = oc
	response.write( Midetalle( 0,2) )
	rs.Fields("producto01"           ) = Midetalle( 0,0)
	rs.Fields("cantidad01"           ) = Midetalle( 0,1)
	rs.Fields("descuento01"          ) = Midetalle( 0,2)
	rs.Fields("producto02"           ) = Midetalle( 1,0)
	rs.Fields("cantidad02"           ) = Midetalle( 1,1)
	rs.Fields("descuento02"          ) = Midetalle( 1,2)
	rs.Fields("producto03"           ) = Midetalle( 2,0)
	rs.Fields("cantidad03"           ) = Midetalle( 2,1)
	rs.Fields("descuento03"          ) = Midetalle( 2,2)
	rs.Fields("producto04"           ) = Midetalle( 3,0)
	rs.Fields("cantidad04"           ) = Midetalle( 3,1)
	rs.Fields("descuento04"          ) = Midetalle( 3,2)
	rs.Fields("producto05"           ) = Midetalle( 4,0)
	rs.Fields("cantidad05"           ) = Midetalle( 4,1)
	rs.Fields("descuento05"          ) = Midetalle( 4,2)
	rs.Fields("producto06"           ) = Midetalle( 5,0)
	rs.Fields("cantidad06"           ) = Midetalle( 5,1)
	rs.Fields("descuento06"          ) = Midetalle( 5,2)
	rs.Fields("producto07"           ) = Midetalle( 6,0)
	rs.Fields("cantidad07"           ) = Midetalle( 6,1)
	rs.Fields("descuento07"          ) = Midetalle( 6,2)
	rs.Fields("producto08"           ) = Midetalle( 7,0)
	rs.Fields("cantidad08"           ) = Midetalle( 7,1)
	rs.Fields("descuento08"          ) = Midetalle( 7,2)
	rs.Fields("producto09"           ) = Midetalle( 8,0)
	rs.Fields("cantidad09"           ) = Midetalle( 8,1)
	rs.Fields("descuento09"          ) = Midetalle( 8,2)
	rs.Fields("producto10"           ) = Midetalle( 9,0)
	rs.Fields("cantidad10"           ) = Midetalle( 9,1)
	rs.Fields("descuento10"          ) = Midetalle( 9,2)
	rs.Fields("producto11"           ) = Midetalle(10,0)
	rs.Fields("cantidad11"           ) = Midetalle(10,1)
	rs.Fields("descuento11"          ) = Midetalle(10,2)
	rs.Fields("producto12"           ) = Midetalle(11,0)
	rs.Fields("cantidad12"           ) = Midetalle(11,1)
	rs.Fields("descuento12"          ) = Midetalle(11,2)
	'rs.fields("obs"                  ) = OBS
	'rs.fields("obs2"                 ) = OBS2
	rs.fields("estado"               ) = "np"
	rs.fields("fechaentrega"		 ) = fc_des
	rs.fields("opcion"				 ) = "VENTA"
	'rs.Update
	rs.Fields("oc"                   ) = right(trim(oc),20)
	rs.fields("obs"                  ) = replace(replace(left(trim(OBS ),100),chr(10)," "),chr(13),".")
	rs.fields("obs2"                 ) = replace(replace(left(trim(OBS2),100),chr(10)," "),chr(13),".")
	rs.fields("empresa"              ) = empresa
	rs.fields("sw_iva"               ) = sw_iva
	rs.fields("tipodocto"            ) = tipodocto
	rs.fields("idbodega"			 ) = idbodega
	rs.Update
	rs.Close

session("echo")=true
response.write("<BR>==========================<BR>Se Genero pedido N:" & nwnota & _
"<BR>==========================<BR>Recibido a las:" & time & _
"<FORM METHOD=POST ACTION='PEDIDO.ASP?nuser=" & nuser & "&empresa=" & empresa & "'>" & _
"<INPUT TYPE='submit' VALUE='Aceptar'></FORM>" )
'-------------------------------------------

'*******************
sql="select * from handheld.dbo.ped_notafolios where idvendedor=" & cint(nuser) & " and idempresa=" & nempresa & ""
Set rs=oConn.execute(Sql)
	sieof=""
	if rs.eof then sieof="si"
	rs.Close
	rs.Open SQL, oConn, 1, 3
	if sieof="si" then
		rs.AddNew
		rs.fields("idvendedor")=cint(nuser)
		rs.fields("idempresa")=nempresa
	end if
	rs.fields("correlativo_notapedido")=nwnota
	rs.Update
	rs.close
'*******************
sql="select top 1 * from handheld.flexline.PDA_CONFIRMA_PED"
Set rs=oConn.execute(Sql)
rs.Close
rs.Open SQL, oConn, 1, 3
rs.AddNew
rs.fields("Numero_pedido")= left(mifecha,4) & right("000" & nuser, 3) & right("0000" & nwnota,4)
rs.fields("fecha")="20" & mifecha
rs.fields("hora")=right("000000" & mihora,6)
rs.Update
rs.close
'************************
end if

set objError = Server.GetLastError()
strErrorLine =  objError.Line

if err<>0 then
	response.write err.description
	response.write strErrorLine
end if
%>
</body>
</html>