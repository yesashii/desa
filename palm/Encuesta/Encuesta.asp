<!--#include Virtual="/includes/conexion.asp"-->
<!--#include Virtual="/includes/forms.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>Encuesta</TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<style type="text/css">
	#noprint {
		display:none;
	}
	body{
		background-color: #F4F4F4;
		font-family: verdana;
		font-size: 12px;
	}
	th{
		font-family: verdana;
		font-size: 12px;
		color: #000033;
		padding: 2px;
		/*border-style: solid;
		border-width: 2px;
		background-color: #000066*/
	}
	td{
		font-family: verdana;
		font-size: 12px;
		padding: 2px;
		/*border-style: solid;
		border-width: 2px;
		padding: 2px;*/
	}
	table{
		border-collapse: collapse;
		/*border: 1px;
		*border-color: #000066;
		background-color: #FFFFFF;*/
	}
</style>
</HEAD>
<!-- ----------------------------------------------------------------------------------------------------- -->
<SCRIPT LANGUAGE="JavaScript">
<!--
function msgcliente(){
	if (document.getElementById('txtcliente') ) {
		var txt = document.getElementById('txtcliente').value;
		if (Length(txt)>0){
			alert(txt);
		}
	}
}//end function msgcliente()
//-----------------------------------------------------------------------------------
function buscasaldomat(producto){

}//end function buscasaldomat
//-----------------------------------------------------------------------------------
function Length(str){return String(str).length;}
//-----------------------------------------------------------------------------------
function Left(str, n){
	if (n <= 0)
	    return "";
	else if (n > String(str).Length)
	    return str;
	else
	    return String(str).substring(0,n);
}
//-----------------------------------------------------------------------------------
function Right(str, n){
    if (n <= 0)
       return "";
    else if (n > String(str).length)
       return str;
    else {
       var iLen = String(str).length;
       return String(str).substring(iLen, iLen - n);
    }
}
//---------------------------------------------------------------------------
function textbox_onChange(objeto){
	if (objeto=='almacen' ){ document.getElementById('idalmacen').value = document.getElementById(objeto).value; };
	if (objeto=='tipoinv' ){ document.getElementById('idtipoinv').value = document.getElementById(objeto).value; };
}//end Funtion textbox_onChange
//-----------------------------------------------------------------------------------
function validainv(){
}
//-----------------------------------------------------------------------------------
function msgbox(msg){window.alert(msg);}
function inputbox(msg){return prompt(msg);}
//-----------------------------------------------------------------------------------
function len(texto){return Length(texto);}
//-----------------------------------------------------------------------------------
function IsNumeric(valor){
	var log=valor.length; var sw="S";
	for (x=0; x<log; x++){ v1=valor.substr(x,1);
		v2 = parseInt(v1);
		if (isNaN(v2)) { sw= "N";}
	}
	if (log==0){sw= "N";};
	if (sw=="S") {return true;} else {return false; }
}
//-----------------------------------------------------------------------------------
function enviarfrm(){
	var okdatos ='1';
	okdatos=validacampo('rut');
	if (okdatos==1){okdatos=validacampo('direccion');}
	if (okdatos==1){okdatos=validacampo('sigla');}
	
	//alert( len(document.getElementById('rut').value) );
	document.getElementById('cl_rut'        ).value=document.getElementById('rut'      ).value;
	document.getElementById('cl_rs'         ).value=document.getElementById('rs'       ).value;
	document.getElementById('cl_direccion'  ).value=document.getElementById('direccion').value;
	document.getElementById('cl_sigla'      ).value=document.getElementById('sigla'    ).value;
	document.getElementById('cl_vendedor'   ).value=rv('vendedor'   );
	document.getElementById('cl_ciudad'     ).value=rv('ciudad'     );
	document.getElementById('cl_cruster'    ).value=rv('cruster'    );
	document.getElementById('cl_canal'      ).value=rv('canal'      );
	
	if(okdatos==1){document.frm2.submit();}
	
}//end enviarfrm
//-----------------------------------------------------------------------------------
function validacampo(objeto){
	if (len(document.getElementById(objeto).value)==0){
		msgbox('Falta '+objeto);
		document.getElementById(objeto).focus();
		 return '0';
	} else {
		 return '1';
	}
	
}//end validacampo
//-----------------------------------------------------------------------------------
function migrilla_onclick(objeto){
	//msgbox(objeto.rowIndex);
	//msgbox(objeto.cells[2].childNodes[0].nodeValue);
	document.getElementById('guardar').value=objeto.rowIndex;
	document.frm3.submit();
}
//-----------------------------------------------------------------------------------
function limpiar(){
	document.getElementById('rut').value='' ;
	document.getElementById('rs'   ).value='' ;
	document.getElementById('direccion'    ).value='' ;
	document.getElementById('sigla'  ).value='' ;
	//document.getElementById('').focus()  ;
	//document.getElementById('').value='' ;
}
//-----------------------------------------------------------------------------------
function validafecha(){
	var fecha=document.getElementById('mfecha').value;
	var vfecha=fecha.split('-');
	var rfecha='';
	if (vfecha.length==1){vfecha=fecha.split('/');}
	if (vfecha.length==3){
		vfecha[1]=Right('00'+vfecha[1],2);
		if (vfecha[2].length==4){
		//dd-mm-aaaa
			vfecha[0]=Right('00'+vfecha[0],2);
			rfecha=''+vfecha[2]+vfecha[1]+vfecha[0]
		}else{
		//aaaa-mm-dd
			vfecha[2]=Right('00'+vfecha[2],2);
			rfecha=''+vfecha[0]+vfecha[1]+vfecha[2]
		}
	}
	document.getElementById('fecha').value=rfecha;
}
//-----------------------------------------------------------------------------------
function rv(idobjeto){//Retorna Valor
	if (document.getElementById(idobjeto)){
		if (document.getElementById(idobjeto).selectedIndex) {
			return document.getElementById(idobjeto).options[document.getElementById(idobjeto).selectedIndex].text;
		} else {
			return document.getElementById(idobjeto).value;
		}
	}else{
		return 'No Encontrado';
	}
}
//-----------------------------------------------------------------------------------

//-->
</SCRIPT>
<%
'-----------------------------------------------------------------------------------------------------------------
'.:: main ::.
rut = recuperavalor("rut")
encuestador=recuperavalor("encuestador")
correlativo  = consultarapida("select isnull(max(idencuesta),0)+1 from sqlserver.desaerp.dbo.TMO_ENCUESTASENC")
if len(rut)>0 then
	if validarut(rut) then
		rs=consultarapida("select razonsocial from serverdesa.BDFlexline.flexline.ctacte where empresa='desa' and tipoctacte='cliente' and codlegal='" & rut & "'")
		if len(rs)>0 then
			'Response.write("Encontrado : " & rs)
			direccion=consultarapida("select direccionenvio from serverdesa.BDFlexline.flexline.ctacte where empresa='desa' and tipoctacte='cliente' and codlegal='" & rut & "'")
			sigla    =consultarapida("select sigla     from serverdesa.BDFlexline.flexline.ctacte where empresa='desa' and tipoctacte='cliente' and codlegal='" & rut & "'")
			nvendedor=consultarapida("select ejecutivo from serverdesa.BDFlexline.flexline.ctacte where empresa='desa' and tipoctacte='cliente' and codlegal='" & rut & "'")
			analisisctacte1=consultarapida("select analisisctacte1 from serverdesa.BDFlexline.flexline.ctacte where empresa='desa' and tipoctacte='cliente' and codlegal='" & rut & "'")
			miciudad=consultarapida("select ciudad from serverdesa.BDFlexline.flexline.ctacte where empresa='desa' and tipoctacte='cliente' and codlegal='" & rut & "'")
	
			'response.write miciudad
		else
			'response.write("Ingresar")
			'Buscar en TMO_CLIENTES
			rs=consultarapida("select sigla from sqlserver.desaerp.dbo.TMO_CLIENTES where idcliente='" & rut & "'")
			if len(rs)>0 then
				sigla=rs
				direccion=consultarapida("select direccion from sqlserver.desaerp.dbo.TMO_CLIENTES where idcliente='" & rut & "'")
				'nvendedor=consultarapida("select direccion from sqlserver.desaerp.dbo.TMO_CLIENTES where idcliente='" & rut & "'")
				'ciudad=consultarapida("select direccion from sqlserver.desaerp.dbo.TMO_CLIENTES where idcliente='" & rut & "'")
				txtcl="Cliente solo en TMO_Clientes"
			else
				txtcl="No se encontro este Cliente en la base de datos, agregar los datos"
			end if
		end if
	else
		txtcl="Rut incorrecto"
	end if
End if





'-----------------------------------------------------------------------------------------------------------------
sub guardar()

	cl_encuestador= recuperavalor("cl_encuestador")
	cl_rut        = recuperavalor("cl_rut"        )
	cl_rs         = recuperavalor("cl_rs"         )
	cl_direccion  = recuperavalor("cl_direccion"  )
	cl_sigla      = recuperavalor("cl_sigla"      )
	cl_vendedor   = recuperavalor("cl_vendedor"   )
	cl_ciudad     = recuperavalor("cl_ciudad"     )
	cl_cruster    = recuperavalor("cl_cruster"    )
	cl_canal      = recuperavalor("cl_canal"      )

	fecha = date()
	hora  = now()
	correlativo  = consultarapida("select isnull(max(idencuesta),0)+1 from sqlserver.desaerp.dbo.TMO_ENCUESTASENC")
	idencuestador=consultarapida("select idencuestador from sqlserver.desaerp.dbo.TMO_ENCUESTADORES")
	idcluster    =consultarapida("select idcluster from sqlserver.desaerp.dbo.tmo_clusters where nombre='" & trim(cl_cruster) & "'")
	idciudad     =consultarapida("select idciudad  from sqlserver.desaerp.dbo.DIM_ciudades where nombre='" & trim(cl_ciudad)  & "'")
	idcanal      =consultarapida("select idcanal   from sqlserver.desaerp.dbo.DIM_canales  where sigla='"  & trim(cl_canal)   & "'")
	etiquetasvino=0
	etiquetasdesa=0

	sql="select top 1 * FROM sqlserver.desaerp.dbo.TMO_CLIENTES where idcliente='" & trim(cl_rut) & "' "
	Set rs=oConn.execute(Sql)
	if rs.eof then
		rs.Close
		rs.Open SQL, oConn, 1, 3
		rs.AddNew
		rs.fields("idcliente") = trim(cl_rut)
		rs.fields("sigla"    ) = cl_sigla
		rs.fields("direccion") = cl_direccion
		rs.fields("idciudad" ) = idciudad 
		rs.fields("idcluster") = idcluster
		rs.fields("idcanal"  ) = idcanal
		rs.Update
		rs.Close
	end if
	':: guarda cliente - fin

	'response.write("<BR>" & cl_rut)
	sql="select top 1 * FROM sqlserver.desaerp.dbo.TMO_ENCUESTASENC "
	Set rs=oConn.execute(Sql)
	rs.Close
	rs.Open SQL, oConn, 1, 3
	rs.AddNew
		rs.fields("idencuesta"   ) = correlativo
		rs.fields("idcliente"    ) = trim(cl_rut)
		rs.fields("idencuestador") =idencuestador
		rs.fields("idcluster"    ) =idcluster
		rs.fields("fecha"        ) =fecha
		rs.fields("hora"         ) =hora
		rs.fields("etiquetasvino") =etiquetasvino
		rs.fields("etiquetasdesa") =etiquetasdesa
	rs.Update
	rs.Close

'.:: detalle ::.
	sql="select * FROM sqlserver.desaerp.dbo.TMO_ENCUESTASDET where idencuesta='" & correlativo & "' "
	Set rs=oConn.execute(Sql)
	rs.Close
	rs.Open SQL, oConn, 1, 3
	for each campo in request.Form
		'Listbox
		if left(campo,3)="cb_" then
			'response.write("<BR>" & campo & " : " & request.form(campo))
			rs.AddNew
			rs.fields("idencuesta" ) = correlativo
			rs.fields("idparametro") =cint(right(campo,3))
			rs.fields("iditem"     ) =request.form(campo)
			rs.Update
		end if
		'CheckListBox
		if left(campo,4)="clb_" then
			rs.AddNew
			rs.fields("idencuesta" ) = correlativo
			rs.fields("idparametro") =cint(left(right(campo,6),3))
			rs.fields("iditem"     ) =cint(right(campo,3))
			rs.Update
		end if	
	next
	rs.Close
	%><CENTER><BR><BR><BR>
	<FORM METHOD=POST ACTION="">
	<TABLE border="1" bgcolor="#FFFFFF">
	<TR>
		<TD>Informe</TD>
	</TR>
	<TR>
		<TD><B>Encuesta <%=correlativo%></B><BR>Los datos se grabaron</TD>
	</TR>
	</TABLE><BR>
	<INPUT TYPE="submit" Value="Nueva Encuesta">
	</FORM>
	</CENTER><%
end sub 'guardar
'-----------------------------------------------------------------------------------------------------------------
function validarut(codlegal)
	validarut=true
end function
'-----------------------------------------------------------------------------------------------------------------
sub encuesta()
%>


<FONT SIZE="2" face="verdana" COLOR="#003366"><B>Cliente&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</B>.</FONT><BR>
<DIV style="width:600px;margin: 10px" align="left">
<FORM METHOD=POST ACTION="">

<TABLE>
<TR style="">
	<TD></TD>
	<TH>Corretativo</TH>
	<TD><%=correlativo%></TD>
</TR>
<TR>
	<TD></TD>
	<TH>Fecha</TH>
	<TD><%= now %></TD>
</TR>
<TR>
	<TD></TD>
	<TH>Encuestador</TH>
	<TD><% 
		sql="select * from sqlserver.desaerp.dbo.TMO_ENCUESTADORES order by nombre"
		call milistbox(sql,"encuestador", midefault)
	%></TD>
</TR>
<TR>
	<TD></TD>
	<TH>. </TH>
	<TD></TD>
</TR>
<TR>
	<TD></TD>
	<TH>Rut</TH>
	<TD><INPUT TYPE="text" NAME="rut" id="rut" size="15" value="<%=rut%>"><INPUT TYPE="submit" value="Validar"><INPUT TYPE="button" value="Limpiar" onclick="limpiar()"></TD>
</TR>
<TR style="visibility:hidden; display:none">
	<TD></TD>
	<TH>Razon Social</TH>
	<TD><INPUT TYPE="text" NAME="rs" id="rs" size="60"  value="<%=rs%>"></TD>
</TR>
<TR style="visibility:hidden; display:none">
	<TD></TD>
	<TH>Direccion</TH>
	<TD><INPUT TYPE="text" NAME="direccion" id="direccion" size="60"  value="<%=direccion%>"></TD>
</TR>
<TR>
	<TD></TD>
	<TH>Sigla</TH>
	<TD><INPUT TYPE="text" NAME="sigla" id="sigla" value="<%=sigla%>"></TD>
</TR>
<TR style="visibility:hidden; display:none">
	<TD></TD>
	<TH>Vendedor</TH>
	<TD><% 
		sql="select nombre from sqlserver.desaerp.dbo.DIM_vendedores where idempresa='1' and sw_palmvigente='S' order by nombre"
		call milistbox(sql,"vendedor", nvendedor)
		%>
	</TD>
</TR>

<TR style="visibility:hidden; display:none">
	<TD></TD>
	<TH>Ciudad</TH>
	<TD>
		<% 
		sql="select nombre from sqlserver.desaerp.dbo.DIM_ciudades order by nombre"
		call milistbox(sql,"ciudad", miciudad)
		%>
	</TD>
</TR>
<TR>
	<TD></TD>
	<TH>Cruster</TH>
	<TD>
		<% 
		sql="select nombre from sqlserver.desaerp.dbo.TMO_Clusters order by nombre"
		call milistbox(sql,"cruster", midefault)
		%>
	</TD>
</TR>
<TR>
	<TD></TD>
	<TH>Canal</TH>
	<TD>
		<% 
		sql="select upper(sigla), sigla from sqlserver.desaerp.dbo.dim_canales order by sigla"
		call milistbox(sql,"canal", analisisctacte1)
		%>
	</TD>
</TR>
<INPUT TYPE="hidden" id="txtcliente" name="txtcliente" value="<%=txtcl%>">
</TABLE>
</FORM>
</DIV>
<HR>
<FONT SIZE="2" face="verdana" COLOR="#003366"><B>Parametros&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</B>.</FONT>
<DIV  style="width:600px;margin: 10px" align="left">
<FORM METHOD=POST ACTION="" id="frm2" name="frm2">
<INPUT TYPE="hidden" id="cl_encuestador" name="cl_encuestador" value="<%=encuestador%>">
<INPUT TYPE="hidden" id="cl_rut"         name="cl_rut"        >
<INPUT TYPE="hidden" id="cl_rs"          name="cl_rs"         >
<INPUT TYPE="hidden" id="cl_direccion"   name="cl_direccion"  >
<INPUT TYPE="hidden" id="cl_sigla"       name="cl_sigla"      >
<INPUT TYPE="hidden" id="cl_vendedor"    name="cl_vendedor"   >
<INPUT TYPE="hidden" id="cl_ciudad"      name="cl_ciudad"     >
<INPUT TYPE="hidden" id="cl_cruster"     name="cl_cruster"    >
<INPUT TYPE="hidden" id="cl_canal"       name="cl_canal"      >
<TABLE >
<%
	sql="SELECT  idparametro, nombre, seleccionmultiple FROM sqlserver.desaerp.dbo.TMO_PARAMETROSENC order by idparametro"
	Set rs1=oConn.execute(sql)
	do until rs1.eof
		%>
		<TR>
		<TD></TD>
		<TH><%=rs1("nombre")%></TH>
		<TD><%
			sql="select iditem, nombre, idclasificador from sqlserver.desaerp.dbo.TMO_PARAMETROSDET where idparametro='" & rs1("idparametro") & "' ORDER BY idparametro, idclasificador, nombre"
			if rs1("seleccionmultiple")=0 then
				call milistbox(sql    ,"cb_"  & right("000" & rs1("idparametro"),3), midefault)
			else
				call checklistbox4(sql,"clb_" & right("000" & rs1("idparametro"),3), midefault)
			end if
		%></TD>
		</TR>
		<%
		rs1.movenext
	loop
%>

</TABLE>
<HR>
<CENTER><INPUT TYPE="button" value="Grabar" onclick="enviarfrm()"></CENTER>
</FORM>
</DIV>
<%
end sub 'encuesta()
'-----------------------------------------------------------------------------------------------------------------

%>
<BODY onload="msgcliente()">
<CENTER>
<%
	if len(recuperavalor("cb_001"))=0 then
		call encuesta()
	else
		call guardar()
	end if
	
%>

</CENTER>
</BODY>
</HTML>
