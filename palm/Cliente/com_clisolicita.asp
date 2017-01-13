<!--#include Virtual="/includes/conexion.asp"-->
<!--#include Virtual="/includes/forms2.asp"-->
<% 
'sub main()
	codlegal           =recuperavalor("codlegal"           )
	ctacte             =recuperavalor("ctacte"             )
	empresa            =recuperavalor("empresa"            )
	guardar            =recuperavalor("guardar"            )
	limitecredito      =recuperavalor("limitecredito"      )
	condpago           =recuperavalor("condpago"           )
	idempresa          =recuperavalor("idempresa"          )
	idvendedor         =recuperavalor("idvendedor"         )
	ejecutivo          =recuperavalor("ejecutivo"          )
	limitecredito_tiene=recuperavalor("limitecredito_tiene")
	idcondpago_tiene   =recuperavalor("idcondpago_tiene"   )

'	response.write("idcondpago_tiene : " & idcondpago_tiene)
	
	idvendedor         =cint(idvendedor)

	'if len(codlegal)=0 then codlegal="13714116-7" 'borrar
	if len(empresa) =0 then empresa="DESA" 
	
	sqlEM = "select idempresa from dim_empresas where nombre ='" & empresa & "'"
	
	set rsIdEM = oConn.execute(sqlEM)
		idempresa=rsIdEM(0)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> COM_CLISOLICITA <%= idempresa & ": " & empresa %></TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<style type="text/css">
	<!--
	body {
		background-color: #FFFFFF;
		font-family: Arial;
		font-size: 12px;
	}
	td {
		font-size: 12px;
		border-collapse: collapse;
	}
	th {
		background-color: #000033;
		font-family: Arial;
		font-size: 12;
		color: #FFFFFF;
	}
	table {
		border: 0;
		border-width:0;
		border-collapse: collapse;
	}
	-->
</style>
</HEAD>

<BODY>
<CENTER>
<%		
		

	if len(guardar)=0 then
		call principal() 
	else
		call guardardatos()
	end if
'end sub main
'-------------------------------------------------------------------------
sub guardardatos()
	'response.write("<BR>" & codlegal     )
	'response.write("<BR>" & empresa      )
	'response.write("<BR>" & guardar      )
	'response.write("<BR>" & condpago     )
	'response.write("<BR>" & limitecredito)

	response.write(idempresa)
	idsucursal=cint(replace(ctacte,codlegal & " ",""))
	sw5_aprobacionlc="P"
	sw6_aprobacioncp="P"
	sw5_nomusuario=""
	sw5_fecha=0
	sw6_nomusuario=""
	sw6_fecha=0
	if limitecredito_tiene=limitecredito then
		sw5_aprobacionlc="A"
		sw5_nomusuario="ROOT"
		sw5_fecha=year(date()) & right("00" & month(date()),2) & right("00" & day(date()),2)
	end if
	if idcondpago_tiene   =condpago      then 
		sw6_aprobacioncp="A"
		sw6_nomusuario="ROOT"
		sw6_fecha=year(date()) & right("00" & month(date()),2) & right("00" & day(date()),2)
	End if

	sql="SELECT top 1 idvendedor FROM sqlserver.desaerp.dbo.DIM_VENDEDORES WHERE (nombre = '" & ejecutivo & "') AND idempresa = " & idempresa
	'response.write(sql)
	set rs=oConn.execute(sql)
	idvendedor =rs(0)
	
'	sql="select top (1) T1.NEMOTECNICO " & _
'	"from serverdesa.BDFlexline.flexline.ctacte T0 " & _
'	"inner join serverdesa.BDFlexline.flexline.GEN_TABCOD T1 ON T0.CondPago=T1.DESCRIPCION " & _
'	"where T0.empresa='" & empresa & "' and T0.codlegal='" & codlegal & "' and T0.tipoctacte='cliente' " & _
'	"and (T1.EMPRESA = 'desa') AND (T1.TIPO = 'GEN_CPAGO_PALM') and T1.nemotecnico not in ('5','11','13','8','3') "
'	set rs=oConn.execute(sql)
'	'response.write("<BR>" & sql)
'	idcondpago_tiene=rs(0)
	
	sql="select isnull(max(idsolicitud),0) as ult from sqlserver.desaerp.dbo.COM_CLISOLICITA"
	set rs=oConn.execute(sql)
	idsolicitud=cdbl(rs("ult"))+1
	response.write("<FONT SIZE='3'><B>Solicitud Nro. " & idsolicitud & "</B></FONT>")
	response.flush()
	
	sql="select * from sqlserver.desaerp.dbo.COM_CLISOLICITA where idsolicitud=" & idsolicitud
	'set rs=oConn.execute(sql) 
	'response.write("<BR>" & idsolicitud)
	'response.write("<BR>" & sql)
	'exit sub 
	rs.close
	rs.Open SQL, oConn, 1, 3
	rs.addnew
	
	rs.fields("idempresa"          )=idempresa
	rs.fields("idsolicitud"        )=idsolicitud
	rs.fields("idcliente"          )=codlegal
	rs.fields("idsucursal"         )=idsucursal          '
	rs.fields("fechasolicitud"     )=year(date()) & right("00" & month(date()),2) & right("00" & day(date()),2)
	rs.fields("idvendedor"         )=idvendedor
	rs.fields("limitecredito_tiene")=limitecredito_tiene
	rs.fields("idcondpago_tiene"   )=idcondpago_tiene
	rs.fields("limitecredito_sol"  )=limitecredito
	rs.fields("idcondpago_sol"     )=condpago
	rs.fields("limitecredito_apr"  )=0
	rs.fields("idcondpago_apr"     )=0
	rs.fields("sw1_informedicom"   )="N"
	rs.fields("sw1_fecha"          )=0
	rs.fields("sw1_nomusuario"     )=""
	rs.fields("sw1_observacion"    )=""
	rs.fields("sw2_autorizadicom"  )="N"
	rs.fields("sw2_fecha"          )=0
	rs.fields("sw2_nomusuario"     )=""
	rs.fields("sw2_observacion"    )=""
	rs.fields("sw3_ivas"           )="N"
	rs.fields("sw3_fecha"          )=0
	rs.fields("sw3_nomusuario"     )=""
	rs.fields("sw3_observacion"    )=""
	rs.fields("sw4_patente"        )="N"
	rs.fields("sw4_fecha"          )=0
	rs.fields("sw4_nomusuario"     )=""
	rs.fields("sw4_observacion"    )=""
	rs.fields("sw5_aprobacionlc"   )=sw5_aprobacionlc
	rs.fields("sw5_fecha"          )=sw5_fecha
	rs.fields("sw5_nomusuario"     )=sw5_nomusuario
	rs.fields("sw5_observacion"    )=""
	rs.fields("sw6_aprobacioncp"   )=sw6_aprobacioncp
	rs.fields("sw6_fecha"          )=sw6_fecha
	rs.fields("sw6_nomusuario"     )=sw6_nomusuario
	rs.fields("sw6_observacion"    )=""
	rs.fields("sw7_refbancaria"    )="N"
	rs.fields("sw7_fecha"          )=0
	rs.fields("sw7_nomusuario"     )=""
	rs.fields("sw7_observacion"    )=""
	
	rs.update
	rs.close
	response.write("<FONT SIZE='2'><BR>La solicitud se guardo Correctamente</FONT>")
	%>
	<BR><BR><INPUT TYPE="button" value="Aceptar" onclick="history.back();history.back();history.back()">
	<%
	''end if
end sub 'guardardatos()
'-------------------------------------------------------------------------
sub principal()
	sql="select * " & _
	"from serverdesa.BDFlexline.flexline.ctacte " & _
	"where empresa='" & empresa & "' and codlegal='" & codlegal & "' and tipoctacte='cliente' " & _
	"order by ctacte"
	set rs=oConn.execute(sql)
	if rs.eof then
		%><CENTER><B>Error datos Cliente</B></CENTER><%
		exit sub
	end if
	
	Sql2="SELECT NEMOTECNICO, DESCRIPCION "&_
		"FROM serverdesa.BDFlexline.flexline.GEN_TABCOD "&_
		"WHERE (EMPRESA = '"& empresa &"') AND (TIPO = 'GEN_CPAGO_PALM') " & _
		"and nemotecnico not in ('5','11','13','8','3')"
	nombre2 ="condpago"
	sqlidpago="select NEMOTECNICO from serverdesa.BDFlexline.flexline.GEN_TABCOD " & _
		"WHERE (EMPRESA = '"& empresa &"') AND (TIPO = 'GEN_CPAGO_PALM') and (DESCRIPCION='" & rs("condpago") & "')"
	default2=consultarapida(sqlidpago)
	
	%>
	<FORM METHOD=POST ACTION="">
	<TABLE width="300px">
	<TR>
		<TH>Solicita cambio credito</TH>
	</TR>
	<TR>
		<TD>
			<TABLE>
			<TR>
				<TD>Cliente</TD>
				<TD><%=pc(rs("codlegal"),2)%></TD>
			</TR>
			<TR>
				<TD>Razon Social</TD>
				<TD><%=pc(rs("razonsocial"),2)%></TD>
			</TR>
			<TR>
				<TD>Ejecutivo</TD>
				<TD><%=pc(idvendedor,2)%></TD>
			</TR>
			<TR>
				<TD>Fecha</TD>
				<TD><%=pc(date(),2)%></TD>
			</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD>
			<TABLE width="100%">
			<TR>
				<TH colspan="2" align="center">Limite Credito</TH>
			</TR>
			<TR>
				<TD>Actual</TD>
				<TD><%=pc(formatnumber(rs("limitecredito"),0),2)%></TD>
			</TR>
			<TR>
				<TD>Nueva</TD>
				<TD><INPUT TYPE="text" NAME="limitecredito" value="<%=rs("limitecredito")%>"></TD>
			</TR>
			<TR>
				<TH colspan="2" align="center">Condicion Pago</TH>
			</TR>
			<TR>
				<TD>Actual</TD>
				<TD><%=pc(rs("condpago"),2)%></TD>
			</TR>
			<TR>
				<TD>Nueva</TD>
				<TD><% call milistbox(sql2, nombre2, default2) %></TD>
			</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD align="center">
<%
'	if len(idempresa)=0 then idempresa=1
	idcondpago_tiene =default2 
%>
			<INPUT TYPE="hidden" value="<%=empresa            %>" name="empresa"            >
			<INPUT TYPE="hidden" value="<%=idempresa          %>" name="idempresa"          >
			<INPUT TYPE="hidden" value="<%=rs("codlegal")     %>" name="codlegal"           >
			<INPUT TYPE="hidden" value="<%=rs("ctacte")       %>" name="ctacte"             >
			<INPUT TYPE="hidden" value="<%=idvendedor         %>" name="idvendedor"         >
			<INPUT TYPE="hidden" value="<%=rs("ejecutivo")    %>" name="ejecutivo"          >
			<INPUT TYPE="hidden" value="<%=rs("limitecredito")%>" name="limitecredito_tiene">
			<INPUT TYPE="hidden" value="<%=idcondpago_tiene   %>" name="idcondpago_tiene"   >
			<INPUT TYPE="button" Value="<< volver" onclick="history.back()">
			<INPUT TYPE="submit" Value="Solicitar Cambio" name="guardar">
		</TD>
	</TR>
	</TABLE>
	</FORM>
	<HR>
	<FORM METHOD=POST ACTION="editabanco.asp">
	<INPUT TYPE="hidden" name="cliente" id="cliente" value="<%=rs("ctacte")%>">
	<INPUT TYPE="submit" value="Cambiar Datos Bancarios">
	</FORM>
	<%
'	response.write("idcondpago_tiene : " & idcondpago_tiene)
end sub 'principal()
'-------------------------------------------------------------------------
%>
</CENTER>
</BODY>
</HTML>