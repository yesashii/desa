<%%>
<html>
<head>
<title>Traspaso</title>
</head>
<body bgcolor="#FFFFFF">
<%
nuser=right("0" & request.cookies("pdausr"),3)
'if session("echo")= true then 
'>
'<p align="center"><b>El Pedido ya fue ingresado</b></p>
'<form method="post" action="pedido.asp?nuser=<%= nuser >">
' align="center"><input type="submit" value="Salir"></p></form>
'<
'else
'========================================================================================
'=============      CAPTURA PEDIDOS PDA    -    version : 1.0 BETA         ==============
'========================================================================================
OC =trim(cstr(request.form("oc"  )))
if len(oc )=0 then oc="0000"
OBS=trim(cstr(request.form("obs" )))
if len(OBS)=0 then OBS = " "
fc_des = request.Form("fE")
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

':::::::::::::::::: conexion :::::::::::::::::
private tipodoc, mitotal, oConn, rs, sql, rs1, SQL2, oConn2
Set oConn  = server.createobject("ADODB.Connection")
oConn.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=todo;User Id=sa;Password=desakey;"

'********************************************************
'********************************************************	
vend=request.QueryString("vend")
SQL2="SELECT vendedor, notap FROM dbo.nota_vtaant WHERE vendedor =N'" & vend & "'"
set rs1=oConn.execute(SQL2)
	if rs1.eof or rs1.bof  then
		nota=0
	else
		nota=rs1.fields("notap")
	end if
	rs1.close
nwnota= (Cdbl(nota) + 1)

'*********************************************************
'*********************************************************
	sql="select * FROM todo.flexline.FX_PEDIDO_PDA FX_PEDIDO_PDA"
	Set rs=oConn.execute(Sql)
	rs.Close
	rs.Open SQL, oConn, 1, 3
	rs.AddNew
	rs.Fields("vendedor"             ) = request.querystring("vend")
	Mifecha=right(year(date),2) & right("00" & month(date),2) & right("00" & day(date),2)
	mihora=right("00" & hour(time),2) & right("00" & minute(time),2) & right("00" & second(time),2)
	rs.Fields("fecha"                ) = mifecha
	rs.Fields("hora"                 ) = mihora
	rs.Fields("cliente"              ) = request.querystring("cliente")
	rs.Fields("nota"                 ) = nwnota
	rs.Fields("oc"                   ) = oc
	rs.Fields("producto01"           ) = Midetalle(0,0)
	rs.Fields("cantidad01"           ) = Midetalle(0,1)
	rs.Fields("descuento01"          ) = Midetalle(0,2)
	rs.Fields("producto02"           ) = Midetalle(1,0)
	rs.Fields("cantidad02"           ) = Midetalle(1,1)
	rs.Fields("descuento02"          ) = Midetalle(1,2)
	rs.Fields("producto03"           ) = Midetalle(2,0)
	rs.Fields("cantidad03"           ) = Midetalle(2,1)
	rs.Fields("descuento03"          ) = Midetalle(2,2)
	rs.Fields("producto04"           ) = Midetalle(3,0)
	rs.Fields("cantidad04"           ) = Midetalle(3,1)
	rs.Fields("descuento04"          ) = Midetalle(3,2)
	rs.Fields("producto05"           ) = Midetalle(4,0)
	rs.Fields("cantidad05"           ) = Midetalle(4,1)
	rs.Fields("descuento05"          ) = Midetalle(4,2)
	rs.Fields("producto06"           ) = Midetalle(5,0)
	rs.Fields("cantidad06"           ) = Midetalle(5,1)
	rs.Fields("descuento06"          ) = Midetalle(5,2)
	rs.Fields("producto07"           ) = Midetalle(6,0)
	rs.Fields("cantidad07"           ) = Midetalle(6,1)
	rs.Fields("descuento07"          ) = Midetalle(6,2)
	rs.Fields("producto08"           ) = Midetalle(7,0)
	rs.Fields("cantidad08"           ) = Midetalle(7,1)
	rs.Fields("descuento08"          ) = Midetalle(7,2)
	rs.Fields("producto09"           ) = Midetalle(8,0)
	rs.Fields("cantidad09"           ) = Midetalle(8,1)
	rs.Fields("descuento09"          ) = Midetalle(8,2)
	rs.Fields("producto10"           ) = Midetalle(9,0)
	rs.Fields("cantidad10"           ) = Midetalle(9,1)
	rs.Fields("descuento10"          ) = Midetalle(9,2)
	rs.Fields("producto11"           ) = Midetalle(10,0)
	rs.Fields("cantidad11"           ) = Midetalle(10,1)
	rs.Fields("descuento11"          ) = Midetalle(10,2)
	rs.Fields("producto12"           ) = Midetalle(11,0)
	rs.Fields("cantidad12"           ) = Midetalle(11,1)
	rs.Fields("descuento12"          ) = Midetalle(11,2)
	rs.fields("obs"                  ) = OBS
	rs.fields("estado"               ) = "np"
	rs.fields("fechaentrega"		 ) = fc_des
	rs.fields("opcion"				 ) = "VENTA"
	rs.Update
	rs.Close

session("echo")=true
response.write("<BR>==========================<BR>Se Genero pedido N:" & nwnota & _
"<BR>==========================<BR>Recibido a las:" & time & _
"<FORM METHOD=POST ACTION='PEDIDO.ASP?nuser=" & nuser & "'>" & _
"<INPUT TYPE='submit' VALUE='Aceptar'></FORM>" )
'-------------------------------------------
'end if
%>
</body>
</html>