<!--#include Virtual="/includes/conexion.asp"-->
<%Public hanaConn
Set hanaConn = server.createobject("ADODB.Connection")
hanaConn.ConnectionTimeOut = 0
hanaConn.CommandTimeout = 0
hanaConn.open "DRIVER={HDBODBC};SERVERNODE=HANAB1:30015;DATABASE=HIS_DESA;UID=SYSTEM;PWD=Passw0rd"
%>
<HEAD>
<TITLE>Incidencias</TITLE>
<style type="text/css">
	body{font-family: verdana;}
	td{font-size: 14px;}
	th{
	font-size: 14px;
	}
</style>
</HEAD>
<body>
<%
'---------------------------------------------------------------------------------------------
'Private sub Main()
	'on error resume next
	set r=response
	
	documento=recuperavalor("factura"  )
	empresa  =recuperavalor("empresa"  )
	tipodocto=recuperavalor("tipodocto")
	'base     =recuperavalor("base"     )
	base="BDFlexline"
	
	if len(empresa  )=0 then empresa  ="DESA"
	if len(tipodocto)=0 then tipodocto="FACT. AFECTA ELEC" '"FACTURAS PALM"
	if len(documento)=0 then documento=recuperavalor("documento")

	if UCASE(empresa)="DESA"  then idempresa=1
	if UCASE(empresa)="LACAV" then idempresa=4

	call DatosDocumento()
'End Private sub Main()
'---------------------------------------------------------------------------------------------
private sub DatosDocumento() 
	r.flush()
	'on error resume next
	dim matriz(10)
	matriz( 1)="Nombre. .   |RAZONSOCIAL"
	matriz( 2)="id Cliente  |CTACTE, SIGLA, FECHAEMISION"
	matriz( 3)="Direccion   |DIRECCIONDESPACHO , COMUNA"
	matriz( 4)="Cond Pago   |CONDICIONPAGO"
	matriz( 5)="Vend Cliente|VENDEDORCLIENTE"
	matriz( 6)="Vend Factura|VENDEDORFACTURA"

	miselect=""
	
	for x=1 to 9
		a=split(matriz(x),"|")
		if ubound(a)=1 then
			'r.write "<BR>u(a):" & ubound(a)
	'		r.write "<BR>a(0):" & a(0)
	'		r.write "<BR>a(1):" & a(1)
			miselect=miselect & ", " & a(1)
		end if
	next
	
	miselect= "SELECT VIGENCIA, TOTAL, CORRELATIVO " & miselect
	miselect= miselect + ", EMPRESA, RIGHT('00000'||CAST(DOCNUM AS VARCHAR),10) AS NUMERO, TIPODOCUMENTO"

	sqldatos=" FROM ""HIS_DESA"".""DESA_DOCENC"" " & _
		"WHERE EMPRESA='" & empresa & "' AND DOCNUM='" & documento & "' and TIPODOCUMENTO='" & tipodocto & "'"
	
'	if tipodocto="GUIA DESPACHO ELEC" then 
'	sqldatos=" from serverdesa.BDFlexline.flexline.documento as Documento " & _
'	"inner join serverdesa.BDFlexline.flexline.ctacte as ctacte " & _
'	"on ctacte.empresa=documento.empresa and ctacte.tipoctacte=documento.tipoctacte and documento.cliente=ctacte.ctacte " & _
'	"inner join (select 'desa' as empresa, '" & tipodocto & "' as tipodocto, '" & documento & "' as correlativo, '' as codigopago) documentop " & _
'	"on documentop.empresa=documento.empresa and documento.tipodocto=documentop.tipodocto and documento.correlativo=documentop.correlativo " & _
'		"UNION " & _
'	"SELECT CASE WHEN Vigencia = 'N' THEN 'S' ELSE 'N' END AS VIGENCIA, " & _
'	"TOTAL, CAST([Numero factura] AS NUMERIC(18, 0)) AS CORRELATIVO, Nombre AS RAZONSOCIAL, " & _
'	"CTACTE, SIGLA, [FECHA EMISION] AS FECHA, [Direccion Despacho] AS DIRECCIONENVIO, 'SIN DATO' AS COMUNA, " & _
'	"[Condicion de Pago] AS CODIGOPAGO, [Vendedor Cliente] AS EJECUTIVO, [Vendedor Factura] AS VENDEDOR, " & _
'	"'DESA' AS EMPRESA,  " & _
'	"RIGHT('0000' + [Numero Factura],10) AS NUMERO, 'FACT. AFECTA ELEC' AS TIPODOCTO " & _
'	"FROM SQLSERVER.DataWarehouse.dbo.HANA_FACTURAENC  " & _
'	")ENCAB " & _
'	"where empresa='" & empresa & "' and numero='" & documento & "' and TipoDocto='" & tipodocto & "'"
'	end if

	sql=miselect & sqldatos

'	sql=replace(ucase(sql),"BDFlexline",base)
	'r.write sql
	set rs=hanaConn.execute(sql)

	if rs.eof then 
		%><CENTER>
		<BR><BR><FONT SIZE="3" COLOR="#000066"><B>Sin Informacion de despacho</B></FONT>
		<BR><BR>
		<INPUT TYPE="button" value="<< Atras" onClick="history.back()">
		</CENTER><%
		exit sub
	end if
	vigencia=rs("vigencia")
	mitotal= FormatNumber(cdbl(rs.fields("total")),0)
	correlativo=rs.fields("correlativo")

	dempresa="Distribucion Excelencia S.A."
	if empresa="DESA"      then dempresa="Distribucion Excelencia S.A."
	if empresa="LACAV"     then dempresa="Distribuidora LA CAV Ltda."
	if empresa="UNDURRAGA" then dempresa="Viña Undurraga"

	%><CENTER>
	<TABLE width='400px'>
	<TR>
		<TD><FONT SIZE="2" COLOR="#6E6E6E"><B><%=dempresa%></B></FONT></TD>
		<TD><FONT SIZE="2" COLOR="#6E6E6E"><B><%=documento%></B></FONT></TD>
	</TR>
	<TR>
		<TD colspan="2"><%
':: Datos cliente
	for x=1 to 9
		a=split(matriz(x),"|")
		if ubound(a)=1 then
			'r.write "<BR>u(a):" & ubound(a)
			'res=a(1)
			res=split(a(1),",")
			'respuesta=ubound(res)
			respuesta=""
			for u=0 to ubound(res)
				'resx=split(res(u),".")
				if ucase(trim(res(u)))="RAZONSOCIAL" then
					respuesta=respuesta & "&nbsp;&nbsp;&nbsp;<B>" & rs(trim(cisnull(res(u),""))) & "</B>"
				else
					respuesta=respuesta & "&nbsp;&nbsp;&nbsp;"    & rs(trim(cisnull(res(u),"")))
				end if
			next

			r.write "<BR><FONT COLOR='#6E6E6E'>" & trim(a(0)) & "&nbsp;:</FONT>" & respuesta
			'r.write "<BR>a(1):" & a(1)
		end if
	next
		%></TD>
	</TR>
	</TABLE><BR><%
	if isnumeric(documento) then documentox=cdbl(documento)
'response.write(documentox)
response.flush()
'------:-----
	sql="select * , " & _
	"cast(year(fecha) as nvarchar) + " & _
	"right('00'+cast(month(fecha) as nvarchar),2) + " & _
	"right('00'+cast(  day(fecha) as nvarchar),2)  " & _
	" as cfecha " & _
	", 'road' as idotransporte ,'' as sw_estado, '' as incidencia ,'' as comentario " & _
	"from master2.flexline.fx_roadmaster as r " & _
	"where r.pedido='" & documentox & "' " & _
	"order by r.fecha"
	sqlroad=replace(sql,"cfecha","fechaentrega")

	sqlOT="select " & _
	"E.idotransporte,  " & _
	"	case len(Fecha) when 8 then  " & _
	"		cast(left(fecha,4) as nvarchar)+'-' +right(left(cast(fecha as nvarchar),6),2)+'-' +right(cast(fecha as nvarchar),2)  " & _
	"	else cast(Fecha as nvarchar) end as FechaEntrega,  " & _
	"t.nombre + ' : ' + t.patente as Vehiculo,  " & _
	"isnull((select top 1 IE.sw_estado " & _
	"from sqlserver.desaerp.dbo.TRP_INCIDENCIASENC as IE  " & _
	"where IE.idnrodocumento=D.idfactura and IE.idotransporte=E.idotransporte),'') as sw_estado, " & _
	"isnull((select top 1 comentario " & _
	"from sqlserver.desaerp.dbo.TRP_INCIDENCIASDET as ID inner join sqlserver.desaerp.dbo.TRP_INCIDENCIASENC as IE  " & _
	"on ID.idempresa=IE.idempresa and  IE.idotransporte=ID.idotransporte and IE.idnrodocumento=ID.idnrodocumento " & _
	"where Id.idnrodocumento=D.idfactura and IE.idotransporte=E.idotransporte),'') as Comentario, " & _
	"isnull((select top 1 II.nombre " & _
	"from sqlserver.desaerp.dbo.TRP_INCIDENCIASDET as ID inner join sqlserver.desaerp.dbo.TRP_INCIDENCIASENC as IE  " & _
	"on ID.idempresa=IE.idempresa and  IE.idotransporte=ID.idotransporte and IE.idnrodocumento=ID.idnrodocumento " & _
	"inner join sqlserver.desaerp.dbo.DIM_INCIDENCIAS as II on II.idincidencia=ID.idincidencia " & _
	"where Id.idnrodocumento=D.idfactura and IE.idotransporte=E.idotransporte),'') as Incidencia, " & _
	"ISNULL((SELECT     TOP (1) ncredito_sap FROM [SQLSERVER].DesaERP.dbo.TRP_ORDTRADEVDET WHERE (idfactura = D.idfactura) AND (idempresa = E.idempresa)), null) AS ncredito_sap " & _
	"from sqlserver.desaerp.dbo.TRP_ORDTRAENC as E inner join sqlserver.desaerp.dbo.TRP_ORDTRADET as D " & _
	"on D.idempresa=E.idempresa and  E.idotransporte=D.idotransporte  " & _
	"inner join sqlserver.desaerp.dbo.TRP_TRANSPORTES AS T  " & _
	"ON E.idempresa = T.idempresa AND E.idtransporte = T.idtransporte AND E.idtransportista = T.idtransportista " & _
	"where D.idfactura='" & documentox & "'  and (E.idestado='V')"
	
	
	sqlOT="select " & _
	"E.idotransporte,   " & _
	"	case len(Fecha) when 8 then  " & _
	"		cast(left(fecha,4) as nvarchar)+'-' +right(left(cast(fecha as nvarchar),6),2)+'-' +right(cast(fecha as nvarchar),2)  " & _
	"	else cast(Fecha as nvarchar) end as FechaEntrega,  " & _
	"t.nombre + ' : ' + isnull((select top 1 nombre from sqlserver.desaerp.dbo.TRP_TRANSPORTISTAS where idtransportista=t.idtransportista and idempresa =E.idempresa ),'') as Vehiculo,  " & _
	"isnull((select top 1 AUT.[idautorizacion] FROM [DesaERP].[dbo].[TRP_INCIDENCIASAUT] AS AUT WHERE E.idotransporte=AUT.idotransporte and E.idempresa=AUT.idempresa and E.idtransportista=AUT.idtransportista and E.idtransporte=AUT.idtransporte and AUT.nrofactura='" & documentox & "'),0) as idautorizacion, " & _
	"isnull((select top 1 IE.sw_estado " & _
	"from sqlserver.desaerp.dbo.TRP_INCIDENCIASENC as IE  " & _
	"where IE.idnrodocumento=D.idfactura and IE.idotransporte=E.idotransporte),'') as sw_estado, " & _
	"isnull((select top 1 comentario  " & _
	"from sqlserver.desaerp.dbo.TRP_INCIDENCIASDET as ID inner join sqlserver.desaerp.dbo.TRP_INCIDENCIASENC as IE  " & _
	"on ID.idempresa=IE.idempresa and  IE.idotransporte=ID.idotransporte and IE.idnrodocumento=ID.idnrodocumento " & _
	"where Id.idnrodocumento=D.idfactura and IE.idotransporte=E.idotransporte),'') + " & _
	"isnull((select top 1 ' ('+nomusuario+')' from sqlserver.desaerp.dbo.TRP_INCIDENCIASENC as IE where IE.idnrodocumento=D.idfactura and IE.idotransporte=E.idotransporte),'')as Comentario, " & _
	"isnull((select top 1 II.nombre " & _
	"from sqlserver.desaerp.dbo.TRP_INCIDENCIASDET as ID inner join sqlserver.desaerp.dbo.TRP_INCIDENCIASENC as IE  " & _
	"on ID.idempresa=IE.idempresa and  IE.idotransporte=ID.idotransporte and IE.idnrodocumento=ID.idnrodocumento " & _
	"inner join sqlserver.desaerp.dbo.DIM_INCIDENCIAS as II on II.idincidencia=ID.idincidencia " & _
	"where Id.idnrodocumento=D.idfactura and IE.idotransporte=E.idotransporte),'') as Incidencia, " & _
	"ISNULL((SELECT     TOP (1) ncredito_sap FROM [SQLSERVER].DesaERP.dbo.TRP_ORDTRADEVDET WHERE (idfactura = D.idfactura) AND (idempresa = E.idempresa)), null) AS ncredito_sap " & _
	"from sqlserver.desaerp.dbo.TRP_ORDTRAENC as E inner join sqlserver.desaerp.dbo.TRP_ORDTRADET as D " & _
	"on D.idempresa=E.idempresa and  E.idotransporte=D.idotransporte  " & _
	"inner join sqlserver.desaerp.dbo.TRP_TRANSPORTES AS T  " & _
	"ON E.idempresa = T.idempresa AND E.idtransporte = T.idtransporte AND E.idtransportista = T.idtransportista and E.idsucursal=T.idsucursal " & _
	"where D.idfactura='" & documentox & "'  and (E.idestado='V') AND T.idtransporte not  in ('505','506','507','508','509','510') and T.nombre NOT IN ('RETORNO','PREPARACION LOGISTICOS')"

'response.write(sqlOT)

'	set rs1=oConn.execute(sqlroad) 'roadmaster
'	if 	rs1.eof then 
	set rs1=oConn.execute(sqlOT) 'desaerp

	'sqlinci="select e.idnrodocumento, e.fechaentrega, e.sw_estado, " & _
	'	"isnull(( " & _
	'	"	select top 1 i.nombre " & _ 
	'	"	from sqlserver.desaerp.dbo.TRP_INCIDENCIASDET as d inner join sqlserver.desaerp.dbo.DIM_INCIDENCIAS as i  " & _
	'	"		on d.idincidencia=i.idincidencia  " & _
	'	"	where d.idempresa=e.idempresa and d.idnrodocumento=e.idnrodocumento and d.fechaentrega=e.fechaentrega " & _
	'	"),' ') as incidencia,  " & _
	'	"isnull(( " & _
	'	"	select top 1 D.comentario  " & _
	'	"	from sqlserver.desaerp.dbo.TRP_INCIDENCIASDET as d " & _
	'	"	where d.idempresa=e.idempresa and d.idnrodocumento=e.idnrodocumento and d.fechaentrega=e.fechaentrega " & _
	'	"),' ') as comentario " & _
	'	"from sqlserver.desaerp.dbo.TRP_INCIDENCIASENC as e " & _
	'	"where e.idempresa = " & idempresa & " and e.idnrodocumento = " & documentox & " and e.idotransporte='" & rs1("idotransporte") & "'"


	'set rs2=oConn.execute(sqlinci)
	'response.write(sqlOT)

	fecha1=20030601
	fecha2=20030601

	'if not rs1.eof then fecha1=rs1("fechaentrega")
	'if not rs2.eof then fecha2=rs2("fechaentrega")
%><!-- <BR><BR>Reporte Consolidado -->
<TABLE border="1">
<TR  bgcolor="#99CCFF">
	<TH>OT</TH>
	<TH>Fecha</TH>
	<TH>Vehiculo</TH>
	<TH>Estado</TH>
	<TH>Incidencia</TH>
	<TH>Nota de Cr&eacute;dito</TH>
	<TH>Comentario</TH>
	<TH>Id Aut.</TH>
</TR>
<%
	do until rs1.eof
		varfecha     =""
		varvehiculo  =""
		varestado    =""
		varincidencia=""
		varcomentario=""
		'simovenext=1
		'on error resume next

		'fecha1=rs1("fechaentrega")
		'varfecha     =cxfecha(fecha1)
		varfecha=rs1("fechaentrega")
		varvehiculo  =PC(rs1("vehiculo"),0)
		varaut = rs1("idautorizacion")
		varestado    ="Por Entregar"
		varincidencia="&nbsp;"
		varcomentario="&nbsp;"
		If len(rs1("ncredito_sap")) > 1 Then
		varNCR = "<a href='documento.asp?documento=" & rs1("ncredito_sap") & "&empresa=" & empresa & "'>" & rs1("ncredito_sap") & "</a>"
		Else
		varNCR = "&nbsp;"
		End If
		idotransporte=rs1("idotransporte")
				if rs1("sw_estado")="E" then varestado="Entregado"
				if rs1("sw_estado")="A" then varestado="Anulado"
				if rs1("sw_estado")="N" then varestado="No Entregado"
				if rs1("sw_estado")="R" then varestado="Redespacho"
				varincidencia=pc(rs1("incidencia"),0) & "&nbsp;"
				varcomentario=PC(rs1("comentario"),0) & "&nbsp;"
				%>
				 <TR>
					<TD><%=idotransporte%></TD>
					<TD><%=varfecha     %></TD>
					<TH><%=varvehiculo  %></TH>
					<TD><%=varestado    %></TD>
					<TD><%=varincidencia%></TD>			
					<TD><%= varNCR %></TD>
					<TD><%=varcomentario%></TD>		
					<TD><%=varaut%></TD>						
				</TR> 
				<%
		rs1.movenext
	loop

%></TABLE><%	


	if ucase(vigencia)="A" then
		%><FONT SIZE="2" COLOR="#CC0000"><B>Documento NULO</B></FONT><%
	end if
	if ucase(vigencia)="N" then
		%><FONT SIZE="2" COLOR="#808080"><B>Documento con Nota de credito mercaderia</B></FONT><%
	end if
	%><BR><TABLE>
	<TR>
		<TD><INPUT TYPE="button" value="<< Atras" onClick="history.back()"></TD>
		<TD><INPUT TYPE="button" value="<< Otro Documento" onClick="history.back();history.back();history.back()"></TD>
		<TD></TD>
	</TR>
	</TABLE><%


End sub 'DatosDocumento() 

'---------------------------------------------------------------------------------------------
function cisnull(yvalor,xvalor)
	if isnull(yvalor) then
		cisnull=xvalor
	else
		cisnull=yvalor
	end if
end function
'---------------------------------------------------------------------------------------------
function cxfecha(fecha)
	cxfecha= right(fecha,2) & "-" & right(left(fecha,6),2) & "-" & left(fecha,4)
end function
'---------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------
sub consulta_ant ()
	sql="select * , " & _
	"cast(year(fecha) as nvarchar) + " & _
	"right('00'+cast(month(fecha) as nvarchar),2) + " & _
	"right('00'+cast(  day(fecha) as nvarchar),2)  " & _
	" as cfecha " & _
	"from master2.flexline.fx_roadmaster as r " & _
	"where r.pedido='" & documentox & "' " & _
	"order by r.fecha"
	sqlruta=replace(sql,"cfecha","fechaentrega")
	%><TABLE border="1">
	    <TR bgcolor="#99CCFF">
			<TH>Fecha</TH>
			<TH>Vehiculo</TH>
			<TH>Estado</TH>
			<TH>Incidencia</TH>
			<TH>Comentario</TH>
		</TR><%
	set rs1=oConn.execute(sql)
	
		
	DO until rs1.eof
		cfecha=rs1("cfecha")
		estado="Por Entregar"
		Incidencia="&nbsp;"
		comentario="&nbsp;"
		sql="SELECT * " & _
		"FROM  sqlserver.desaerp.dbo.TRP_INCIDENCIASENC as E " & _
		"WHERE     (E.idempresa = " & idempresa & ") AND (E.idnrodocumento = " & documentox & ") AND (E.fechaentrega = " & cfecha & ")"
		'r.write(sql)
		set rs2=oConn.execute(sql)
		if not rs2.eof then
			'estado=rs2("estado")
			
			'estado=rs2("sw_estado")
			if rs2("sw_estado")="E" then estado="Entregado"
			if rs2("sw_estado")="A" then estado="Anulado"
			if rs2("sw_estado")="N" then estado="No Entregado"
			if rs2("sw_estado")="R" then estado="Redespacho"
			'if lcase(rs2("sw_entregado"))="s" then estado="Entregado"
			
			sql="SELECT     i.nombre, d.comentario, e.sw_estado, e.sw_entregado " & _
			"FROM  sqlserver.desaerp.dbo.TRP_INCIDENCIASDET as D " & _
			"INNER JOIN sqlserver.desaerp.dbo.TRP_INCIDENCIASENC as E " & _
			"   ON D.idempresa = E.idempresa AND D.idnrodocumento = E.idnrodocumento AND D.fechaentrega = E.fechaentrega " & _
			"INNER JOIN sqlserver.desaerp.dbo.DIM_INCIDENCIAS as I " & _
			"   ON D.idincidencia = I.idincidencia " & _
			"WHERE     (D.idempresa = " & idempresa & ") AND (D.idnrodocumento = " & documentox & ") AND (D.fechaentrega = " & cfecha & ")"
			set rs2=oConn.execute(sql)
			if not rs2.eof then
				Incidencia=pc(rs2("nombre"),0)
				comentario=pc(rs2("comentario") & "&nbsp;",0)
			end if

		end if
		'r.write 
		%><TR>
			<TD><%=rs1("fecha")%></TD>
			<TD align="center"><B><%=rs1("vehiculo")%></B></TD>
			<TD><%=estado%></TD>
			<TD><%=Incidencia%></TD>
			<TD><%=comentario%></TD>
		</TR><%
	rs1.movenext
	loop
	%></TABLE><%
	r.flush()
	sql="select inci.* " & _
	"from ( " & _
	"	SELECT cast(D.idnrodocumento as nvarchar)+'_'+cast(D.fechaentrega as nvarchar) as idx , " & _
	"	D.fechaentrega,i.nombre, d.comentario, e.sw_estado, e.sw_entregado " & _
	"	FROM sqlserver.desaerp.dbo.TRP_INCIDENCIASDET as D " & _
	"	INNER JOIN sqlserver.desaerp.dbo.TRP_INCIDENCIASENC as E " & _ 
	"	ON D.idempresa = E.idempresa AND D.idnrodocumento = E.idnrodocumento AND D.fechaentrega = E.fechaentrega " & _
	"	INNER JOIN sqlserver.desaerp.dbo.DIM_INCIDENCIAS as I " & _
	"	ON D.idincidencia = I.idincidencia " & _
	"	WHERE (D.idempresa = " & idempresa & ") AND (D.idnrodocumento = " & documentox & ") " & _
	") as Inci left outer join( " & _
	"	select pedido+'_'+master2.flexline.yyyymmdd(fecha) as idx,* " & _
	"	from master2.Flexline.FX_ROADMASTER " & _
	"	where Pedido=" & documentox & " " & _
	") as Road " & _
	"on inci.idx=road.idx " & _
	"where road.idx is null"
	
	sql="select inci.* " & _
	"from (  " & _
	"	select x.*, isnull(i.nombre,' ') as nombre " & _
	"	from ( " & _
	"		SELECT  " & _
	"			cast(E.idnrodocumento as nvarchar)+'_'+cast(E.fechaentrega as nvarchar) as idx ,  " & _
	"			E.fechaentrega, " & _
	"			isnull(d.comentario,' ') as comentario,  " & _
	"			e.sw_estado,  " & _
	"			e.sw_entregado, " & _
	"			isnull(d.idincidencia,' ') as idincidencia   " & _
	"		FROM sqlserver.desaerp.dbo.TRP_INCIDENCIASDET as D  " & _
	"		right outer JOIN sqlserver.desaerp.dbo.TRP_INCIDENCIASENC as E  " & _
	"			ON cast(D.idempresa as nvarchar)+'_'+cast(D.idnrodocumento as nvarchar)+'_'+cast(D.fechaentrega as nvarchar)= " & _
	"			cast(E.idempresa as nvarchar)+'_'+cast(E.idnrodocumento as nvarchar)+'_'+cast(E.fechaentrega as nvarchar) " & _
	"		WHERE (E.idempresa = " & idempresa & ") AND (E.idnrodocumento = " & documentox & " )  " & _
	"	) as x left outer join sqlserver.desaerp.dbo.DIM_INCIDENCIAS as I " & _
	"	ON x.idincidencia = I.idincidencia " & _
	") as Inci left outer join(  " & _
	"	select  " & _
	"	pedido+'_'+master2.flexline.yyyymmdd(fecha) as idx, " & _
	"	*  " & _
	"	from master2.Flexline.FX_ROADMASTER  " & _
	"	where Pedido=" & documentox & "   " & _
	") as Road on inci.idx=road.idx  " & _
	"where road.idx is null " 
	set rs=oConn.execute(sql)
	'r.write(sql)
	if not rs.eof then
		%><BR><FONT SIZE="2" COLOR="#000033">Incidencias sin ruta</FONT><TABLE border="1">
		<TR bgcolor="#99CCFF">
			<TH>Fecha</TH>
			<TH>Estado</TH>
			<TH>Incidencia</TH>
			<TH>Comentario</TH>
		</TR><%
	end if
	do until rs.eof
		fecha=cxfecha(rs("fechaentrega"))
		incidencia=rs("nombre")  & "&nbsp;"
		comentario=rs("comentario") & "&nbsp;"
		if rs("sw_estado")="E" then estado="Entregado"
		if rs("sw_estado")="A" then estado="Anulado"
		if rs("sw_estado")="N" then estado="No Entregado"
		if rs("sw_estado")="R" then estado="Redespacho"
		%><TR>
			<TD><%=fecha%></TD>
			<TD><%=estado%></TD>
			<TD><%=incidencia%></TD>
			<TD><%=comentario%></TD>
		</TR><%
	rs.movenext
	loop
	%></TABLE><%


end sub
'---------------------------------------------------------------------------------------------

%>
</body>