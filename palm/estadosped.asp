<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>EstadoPed</TITLE>
<META NAME="Generator" CONTENT="www.Sistcomp.cl">
<META NAME="Author" CONTENT="Simon Hernandez">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
</HEAD>
<body bgcolor="#FFFFFF" topmargin="0">
<CENTER>
<%
':::::::::::::::::: conexion :::::::::::::::::
'Response.ContentType = "application/vnd.ms-excel"
Dim tipodoc, mitotal, oConn, rs, sql, oconn1, rs1, Midia, Mimes, Miano, Micolor
DIM BdNombre, BDdireccion, BDFono, BDServer, BDuser, BDpass, BDimagen, bdbase
Dim VUNombre, VUtipousr, VUnum_vend
Set oConn = server.createobject("ADODB.Connection")
Set oConn1 = server.createobject("ADODB.Connection")
oConn.open "provider=Microsoft.jet.OLEDB.4.0; Data Source=" & server.mapPath("base.mdb")
oConn1.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=handheld;User Id=sa;Password=desakey;"

if len(nusuario)<1 then nusuario=request.form("nuser")
if len(nusuario)<1 then nusuario=request.querystring("nuser")
if len(nusuario)<1 then nusuario=request.cookies("fx_nusuario")

if len(Miano)<1 then Miano=request.form("Miano")
if len(Miano)<1 then Miano=request.querystring("Miano")
if len(Miano)<1 then Miano=right(year(date),2)

if len(Midia)<1 then Midia=request.form("Midia")
if len(Midia)<1 then Midia=request.querystring("Midia")
if len(Midia)<1 then Midia=Right("00" & day(date),2)

if len(Mimes)<1 then Mimes=request.form("Mimes")
if len(Mimes)<1 then Mimes=request.querystring("Mimes")
if len(Mimes)<1 then Mimes=Right("00" & month(date),2)
mifechax="20" & miano & mimes & midia
if cdbl(mifechax) > 20080219 then
	%><BR><BR>Por Cambio de sistema ...<BR>
	<FORM METHOD=POST ACTION="estadosped_ra.asp?nuser=<%=nusuario%>&mifecha=<%=mifechax%>">
	<INPUT TYPE="submit" value="Ir al Nuevo Reporte">
	</FORM><%
else
'response.write(mifechax)
Sql="select * from Flexline.FX_vende_man where nu_2='" & nusuario &"'"
'SELECT     *
'FROM         Flexline.FX_vende_man
Set rs=oConn1.execute(Sql)

	Call Titulo()
	Call Mitabla()
end if
'--------------------------------------------------------------------
Sub Titulo()
%>
<TABLE width='100%' border='0' cellspacing='0' cellpadding='0'>
	<TR>
		<TD bgcolor='#000066' align='center'>
			<FONT SIZE='' COLOR='#FFFFFF'>
				<B><%=nusuario & "&nbsp;" & rs.fields("Nombre")%></B>
			</FONT>
		</TD>
		<TD bgcolor='#000066' align='center'>
			<A HREF='Default.asp'>
				<IMG SRC='home.jpg' WIDTH='8' HEIGHT='7' BORDER='0' ALT=''>
			</A>
			<A HREF='Config.asp'>
				<IMG SRC='conf.jpg' WIDTH='8' HEIGHT='7' BORDER='0' ALT=''>
			</A>
		</TD>
	</TR>
</TABLE>
	<FORM METHOD=POST ACTION="estadosped.asp" >
	Dia<SELECT NAME="midia" >
	<%
'lista dias
x=1
do until x > 31
	xDia=right("00"& x,2)
	xselec=""
	if x=cdbl(Midia) then xselec="SELECTED "
	response.write("<OPTION " & xselec & "value=" & xDia & ">" & xDia & "</OPTION>")
x=x+1
loop
	%>
	</SELECT>
	Mes
	<INPUT TYPE="text" NAME="mimes" size="2" value=<%=Mimes%>>&nbsp;
	Año
	<INPUT TYPE="text" NAME="miano" size="2" value=<%=Miano%>>
	<INPUT TYPE="hidden" NAME="nuser" value=<%=nusuario%>>
	<INPUT TYPE="submit" value="Aplicar">
	<HR>
	</FORM>
<%
End Sub 'Titulo()
'--------------------------------------------------------------------
Sub Mitabla()
'on error resume next
%>
<TABLE BORDER="0" CELLSPACING="0">
<TR BGCOLOR="#000066">
	<TD ALIGN=center><B><FONT SIZE="2" COLOR="#FFFFFF" Face="arial">Nota</FONT   ></B></TD>
	<TD ALIGN=center><B><FONT SIZE="2" COLOR="#FFFFFF" Face="arial">Cliente</FONT></B></TD>
	<TD ALIGN=center><B><FONT SIZE="2" COLOR="#FFFFFF" Face="arial">Est</FONT    ></B></TD>
	<TD ALIGN=center><B><FONT SIZE="2" COLOR="#FFFFFF" Face="arial">Detalle</FONT></B></TD>
</TR>
<%
MiFecha=Miano & Mimes & Midia
SQL="SELECT FX_PEDIDO_PDA.nota, FX_PEDIDO_PDA.Cliente, CtaCte.Sigla, CtaCte.RazonSocial, FX_PEDIDO_PDA.estado, FX_PEDIDO_PDA.Nfactura, " & _
"FX_PEDIDO_PDA.estado2, FX_PEDIDO_PDA.DetalleR, FX_PEDIDO_PDA.Notaventa, FX_PEDIDO_PDA.ID, " & _
"FX_PEDIDO_PDA.fechaentrega "&_
"FROM Flexline.CtaCte CtaCte, Flexline.FX_PEDIDO_PDA FX_PEDIDO_PDA " & _
"WHERE FX_PEDIDO_PDA.Cliente = CtaCte.CtaCte AND " & _
"((FX_PEDIDO_PDA.vendedor='" & rs.fields("Nombre") & "') AND " & _
"(FX_PEDIDO_PDA.fecha='" & MiFecha & "') AND " & _
"(FX_PEDIDO_PDA.estado In ('np','OK')) AND (CtaCte.Empresa='desa') AND (CtaCte.TipoCtaCte='cliente')) " & _
"ORDER BY FX_PEDIDO_PDA.nota"
Set rs1=oConn1.execute(Sql)
'response.write(SQL)

dim micolor
Micolor="#FFFF99"

Do until rs1.eof
	Nombre=""
	if Not isnull(rs1.fields("Sigla")) then Nombre=trim(rs1.fields("Sigla"))
	if len(Nombre)=0 then
		Nombre=left(trim(rs1.fields("RazonSocial")),30 )
	else
		Nombre=left("<B>" & Nombre & "</B>" & " " & trim(rs1.fields("RazonSocial")),37 )
	end if

	Estado ="NP"
	if rs1.fields("estado")="OK" then Estado ="PP"
	if Not isnull(rs1.fields("fechaentrega")) then
eDia= right(rs1.fields("fechaentrega"),2)
eMes= mid(rs1.fields("fechaentrega"),5,2)	

fEntrega = cdate(eDia & "/" & eMes &"/"& year(date))
    if fEntrega >= date then 
'response.Write weekdayname(weekday(fEntrega))
         fEntrega = dateadd("d", 1, fEntrega)
	     Detalle = "Entrega " & weekdayname(weekday(fEntrega))
	else
	Detalle="sin detalle"
	end if
else
	Detalle="sin detalle"
end if
	
	if not isnull(rs1.fields("estado2")) then 
		Estado ="<B>" & rs1.fields("estado2") & "</B>"
		Detalle=rs1.fields("DetalleR")
		if left(Detalle,1)="_" then Detalle=right(Detalle,len(Detalle)-1)
	End if
	if not isnull(rs1.fields("Nfactura")) then 'documento.asp?factura=0000548842
		if len(rs1.fields("Nfactura"))>5 then 
			Numero=Right("0000000000" & rs1.fields("Nfactura"),10)
			Detalle="<A HREF='documento.asp?factura=" & Numero & "' " & _
				"style='text-decoration: none' >" & _
				rs1.fields("Nfactura") & "</A>"
			Estado="OK"
		end If
	end if
if len(rs1.fields("Notaventa"))>5 then
MiHREF="HREF='NotaVenta.asp?factura=" & rs1.fields("Notaventa")
else
MiHREF="HREF='buscanota.asp?user=" & rs.fields("Nombre") & "&nota=" & rs1.fields("nota")
End if
	Nota="<A " & MiHREF & "' style='text-decoration: none' >" & rs1.fields("nota") & "</A>"
%>
<TR BGCOLOR="<%=micolor%>">
	<TD ALIGN=center><%=Nota%></TD>
	<TD><%=replace(ProperCase(Nombre)," ","&nbsp;") %></TD>
	<TD ALIGN=center><%=Estado%></TD>
	<TD ALIGN=center><%=Detalle%></TD>
</TR>
<%

if micolor="#FFFF99" then
	micolor="#CCFFFF"
else
	micolor="#FFFF99"
end if 
rs1.movenext
loop
	
%></TABLE><%

call PedidosFono()
call PedidosOtros()
call Cambiafecha()

End Sub 'Mitabla()
'--------------------------------------------------------------------
Sub Cambiafecha()
%>
<HR>
<TABLE>
<TR>
	<TD>
	<%if midia<>"01" then%>
	<FORM METHOD=POST ACTION="">
		<INPUT TYPE="hidden" NAME="mimes" size="2" value=<%=mimes%>>
		<INPUT TYPE="hidden" NAME="miano" size="2" value=<%=miano%>>
		<INPUT TYPE="hidden" NAME="midia" size="2" value=<%= right("00" & cint(midia)-1,2)%>>
		<INPUT TYPE="hidden" NAME="nuser" value=<%=nusuario%>>
		<INPUT TYPE="submit" value="<< Ant">
	</FORM>
	<%End if%>
	</TD>
	<TD>
		D&iacute;a
	</TD>
	<TD>
	<%if midia<>"31" then%>
	<FORM METHOD=POST ACTION="">
		<INPUT TYPE="hidden" NAME="mimes" size="2" value=<%=mimes%>>
		<INPUT TYPE="hidden" NAME="miano" size="2" value=<%=miano%>>
		<INPUT TYPE="hidden" NAME="midia" size="2" value=<%= right("00" & cint(midia)+1,2)%>>
		<INPUT TYPE="hidden" NAME="nuser" value=<%=nusuario%>>
		<INPUT TYPE="submit" value="Sig >>">
	</FORM>
	<%End if%>
	</TD>
</TR>
</TABLE>
<%
End Sub 'Cambiafecha()
'--------------------------------------------------------------------------
Sub PedidosFono()
'on error resume next
MiFecha=Miano & Mimes & Midia


SQL="SELECT FX_PEDIDO_PDA.nota, FX_PEDIDO_PDA.Cliente, CtaCte.Sigla, CtaCte.RazonSocial, FX_PEDIDO_PDA.estado, " & _
"FX_PEDIDO_PDA.NFactura, FX_PEDIDO_PDA.vendedor, CtaCte.Ejecutivo, " & _
"FX_PEDIDO_PDA.estado2, FX_PEDIDO_PDA.DetalleR, FX_PEDIDO_PDA.Notaventa, FX_PEDIDO_PDA.ID " & _
"FROM Flexline.CtaCte CtaCte INNER JOIN " & _
"Flexline.FX_PEDIDO_PDA FX_PEDIDO_PDA ON CtaCte.CtaCte = FX_PEDIDO_PDA.Cliente " & _
"WHERE (FX_PEDIDO_PDA.vendedor LIKE '%call%') AND (FX_PEDIDO_PDA.fecha = '" & MiFecha & "') AND (FX_PEDIDO_PDA.estado IN ('np', 'OK')) AND " & _
"(CtaCte.Empresa = 'desa') AND (CtaCte.TipoCtaCte = 'cliente') AND (CtaCte.Ejecutivo = N'" & rs.fields("Nombre") & "') " & _
"ORDER BY FX_PEDIDO_PDA.nota"

'response.write(sql)
Set rs1=oConn1.execute(Sql)
if rs1.eof then exit sub
%>

<BR><FONT SIZE="2" COLOR="#8C8C8C">Pedidos Call Center</FONT>
<TABLE BORDER="0" CELLSPACING="0">

<TR BGCOLOR="#000066">
	<TD ALIGN=center><B><FONT SIZE="2" COLOR="#FFFFFF" Face="arial">Nota</FONT   ></B></TD>
	<TD ALIGN=center><B><FONT SIZE="2" COLOR="#FFFFFF" Face="arial">Cliente</FONT></B></TD>
	<TD ALIGN=center><B><FONT SIZE="2" COLOR="#FFFFFF" Face="arial">Est</FONT    ></B></TD>
	<TD ALIGN=center><B><FONT SIZE="2" COLOR="#FFFFFF" Face="arial">Detalle</FONT></B></TD>
</TR>
<%

'response.write(sql)
dim micolor
Micolor="#FFFF99"

Do until rs1.eof
	Nombre=""
	if Not isnull(rs1.fields("Sigla")) then Nombre=trim(rs1.fields("Sigla"))
	if len(Nombre)=0 then
		Nombre=left(trim(rs1.fields("RazonSocial")),30 )
	else
		Nombre=left("<B>" & Nombre & "</B>" & " " & trim(rs1.fields("RazonSocial")),37 )
	end if

	Estado ="NP"
	if rs1.fields("estado")="OK" then Estado ="PP"
	Detalle="sin detalle"
	
	if not isnull(rs1.fields("estado2")) then 
		Estado ="<B>" & rs1.fields("estado2") & "</B>"
		Detalle=rs1.fields("DetalleR")
		if left(Detalle,1)="_" then Detalle=right(Detalle,len(Detalle)-1)
	End if
	if not isnull(rs1.fields("Nfactura")) then 'documento.asp?factura=0000548842
		if len(rs1.fields("Nfactura"))>5 then 
			Numero=Right("0000000000" & rs1.fields("Nfactura"),10)
			Detalle="<A HREF='documento.asp?factura=" & Numero & "' " & _
				"style='text-decoration: none' >" & _
				rs1.fields("Nfactura") & "</A>"
			Estado="OK"
		end If
	end if
if len(rs1.fields("Notaventa"))>5 then
MiHREF="HREF='NotaVenta.asp?factura=" & rs1.fields("Notaventa")
else
MiHREF="HREF='buscanota.asp?user=" & rs1.fields("vendedor") & "&nota=" & rs1.fields("nota")
End if
	Nota="<A " & MiHREF & "' style='text-decoration: none' >" & rs1.fields("nota") & "</A>"
%>
<TR BGCOLOR="<%=micolor%>">
	<TD ALIGN=center><%=Nota%></TD>
	<TD><%=replace(ProperCase(Nombre)," ","&nbsp;") %></TD>
	<TD ALIGN=center><%=Estado%></TD>
	<TD ALIGN=center><%=Detalle%></TD>
</TR>
<%

if micolor="#FFFF99" then
	micolor="#CCFFFF"
else
	micolor="#FFFF99"
end if 
rs1.movenext
loop

%>
</TABLE>
<%

End Sub 'PedidosFono()
'--------------------------------------------------------------------
Sub PedidosOtros()
'on error resume next
MiFecha="20" & Miano & "-" & MiMes & "-" & Midia


SQL="SELECT     Documento_EC_1.Numero, Documento_EC_1.TipoDocto, Flexline2.dbo.CtaCte.Sigla, Flexline2.dbo.CtaCte.RazonSocial, " & _
"                      Documento_EC_1.ReferenciaExterna, Flexline2.dbo.TipoDocumento.clase " & _
"FROM         Flexline2.Flexline.Documento_EC Documento_EC_1 INNER JOIN " & _
"                      Flexline2.dbo.TipoDocumento ON Documento_EC_1.Empresa = Flexline2.dbo.TipoDocumento.Empresa AND  " & _
"                      Documento_EC_1.TipoDocto = Flexline2.dbo.TipoDocumento.TipoDocto INNER JOIN " & _
"                      Flexline2.dbo.CtaCte ON Documento_EC_1.Empresa = Flexline2.dbo.CtaCte.Empresa AND  " & _
"                      Documento_EC_1.TipoCtaCte = Flexline2.dbo.CtaCte.TipoCtaCte AND Documento_EC_1.Cliente = Flexline2.dbo.CtaCte.CtaCte " & _
"GROUP BY Documento_EC_1.Empresa, Documento_EC_1.Vendedor, Documento_EC_1.Fecha, Documento_EC_1.TipoDocto,  " & _
"                      Flexline2.dbo.TipoDocumento.Sistema, Flexline2.dbo.TipoDocumento.FactorMonto, Documento_EC_1.ReferenciaExterna, Documento_EC_1.Numero, Flexline2.dbo.TipoDocumento.clase,  " & _
"                      Flexline2.dbo.CtaCte.RazonSocial, Flexline2.dbo.CtaCte.Sigla " & _
"HAVING      (Documento_EC_1.Empresa = 'desa') AND (Documento_EC_1.Vendedor = '" & rs.fields("Nombre") & "') AND (Flexline2.dbo.TipoDocumento.Sistema = 'ventas')  " & _
"                      AND (Flexline2.dbo.TipoDocumento.FactorMonto <> 0) AND (Documento_EC_1.TipoDocto <> 'FACTURAS PALM') AND  " & _
"                      (Documento_EC_1.Fecha = CONVERT(DATETIME, '" & MiFecha & " 00:00:00', 102))"

'response.write(sql)
Set rs1=oConn1.execute(Sql)
if rs1.eof then exit sub
%>

<BR><FONT SIZE="2" COLOR="#8C8C8C">Otros Documentos</FONT>
<TABLE BORDER="0" CELLSPACING="0">

<TR BGCOLOR="#000066">
	<TD ALIGN=center><B><FONT SIZE="2" COLOR="#FFFFFF" Face="arial">Numero</FONT   ></B></TD>
	<TD ALIGN=center><B><FONT SIZE="2" COLOR="#FFFFFF" Face="arial">Cliente</FONT></B></TD>
	<TD ALIGN=center><B><FONT SIZE="2" COLOR="#FFFFFF" Face="arial">Documento</FONT    ></B></TD>
	<TD ALIGN=center><B><FONT SIZE="2" COLOR="#FFFFFF" Face="arial">Referencia</FONT></B></TD>
</TR>
<%

'response.write(sql)
dim micolor
Micolor="#FFFF99"

Do until rs1.eof
	Nombre=""
	if Not isnull(rs1.fields("Sigla")) then Nombre=trim(rs1.fields("Sigla"))
	if len(Nombre)=0 then
		Nombre=left(trim(rs1.fields("RazonSocial")),30 )
	else
		Nombre=left("<B>" & Nombre & "</B>" & " " & trim(rs1.fields("RazonSocial")),37 )
	end if

Nota=rs1.fields("Numero")
MiHREF="HREF='Documento.asp?factura=" & rs1.fields("Numero")
if rs1.fields("clase")="Factura (v)" then Nota="<A " & MiHREF & "' style='text-decoration: none' >" & rs1.fields("Numero") & "</A>"

Detalle=rs1.fields("ReferenciaExterna")
Estado="<B>" & replace(replace(rs1.fields("TipoDocto"),"FACTURA","Fac"),"NOTA DE CREDITO","N.C. ") & "</B>"
%>
<TR BGCOLOR="<%=micolor%>">
	<TD ALIGN=center><%=Nota%></TD>
	<TD><%=replace(ProperCase(Nombre)," ","&nbsp;") %></TD>
	<TD ALIGN=center><%=Estado%></TD>
	<TD ALIGN=center><%=Detalle%></TD>
</TR>
<%

if micolor="#FFFF99" then
	micolor="#CCFFFF"
else
	micolor="#FFFF99"
end if 
rs1.movenext
loop

%>
</TABLE>
<%
End Sub 'PedidosOtros()
'--------------------------------------------------------------------
Function ProperCase(sString) 
  ' Declare the variables 
  Dim sWhiteSpace, bCap, iCharPos, sChar 

  ' Valid white space characters. Add more as necessary
  sWhiteSpace = Chr(32) & Chr(9) & Chr(13) 

  ' Convert the string to lowercase, to start with 
  sString = LCase(sString)

  ' Initialize the caps flag to True 
  bCap = True 
  
  ' Loop through each of the characters in the string 
  For iCharPos = 1 to Len(sString) 
    ' Get one character at a time
    sChar = Mid(sString, iCharPos, 1) 

    ' If caps flag is true, upper case the character 
    If bCap = True Then 
      sChar = UCase(sChar) 
    End If 

    ' Append to the final string 
    ProperCase = ProperCase + sChar 

    ' If the character is a white space, raise the caps flag 
    If InStr(sWhiteSpace, sChar) Then 
      bCap = True 
    Else 
      bCap = False 
    End If 
  Next 
End Function 
'_____________________________________________________________________
%>
<FONT SIZE="1" COLOR="#808080">
	<U>Estado</U><BR>
	NP=No Procesado
	PP=Pedido Prosesado
	OK=Aprobado
	RD=Rezhazado
</FONT>
<FONT SIZE="1" COLOR="#ADADAD"><BR>
Distribuci&oacute;n y Excelencia S.A.<BR>
____________________<BR>
Fono: 4891500
</FONT>
</CENTER>