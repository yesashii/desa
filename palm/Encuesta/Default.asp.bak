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
	var nsem=document.getElementById('semana').value;
	var nper=document.getElementById('periodo').value;
	if (objeto=='almacen' ){ document.getElementById('idalmacen').value = document.getElementById(objeto).value; };
	if (objeto=='tipoinv' ){ document.getElementById('idtipoinv').value = document.getElementById(objeto).value; };
	if (objeto=='semana'  ){ cargafecha(nper,nsem)}
	if (objeto=='periodo' ){ cargafecha(nper,nsem)}
	if (objeto=='ciudad' ){ 
		//alert(document.getElementById('ciudad').value)
		if (document.getElementById('ciudad').value=='Santiago'){
		}else{
			document.getElementById('cluster').value='Ninguno';
		}
	}
	if (objeto=='cluster' ){ 
		//alert(document.getElementById('ciudad').value)
		if (document.getElementById('cluster').value=='Ninguno'){

		}else{
			if (document.getElementById('ciudad').value=='Santiago'){
			}else{
				alert('El cluster solo aplica a santiago');
				document.getElementById('cluster').value='Ninguno';
			}
		}
	}
}//end Funtion textbox_onChange
//-----------------------------------------------------------------------------------
function cargafecha(xperiodo,xsemana){
	for(i=0;i<Msemana.length;i++){
		if (Msemana[i][0]==xperiodo && Msemana[i][1]==xsemana){
			document.getElementById('desde').value=Msemana[i][2];
			document.getElementById('hasta').value=Msemana[i][3];
		};
	};
}
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
<BODY>
<CENTER>
<%
'------------------------------------------------------------------------------------
'.:: sub main  ::.
	
	if len(recuperavalor("sigla"))=0 then
		call encuesta()
	else
		call guardar()
	end if
	

'.:: end sub main ::.
'------------------------------------------------------------------------------------
sub encuesta()
call Jmatrix("Msemana","select periodo, semana, desde, hasta from sqlserver.desaerp.dbo.TMO_semanas")
%><FORM METHOD=POST ACTION="">

<TABLE>
<TR>
	<TH>Encuestador</TH>
	<TD>
	<% 
		sql="select nombre from sqlserver.desaerp.dbo.TMO_ENCUESTADORES order by nombre"
		call milistbox(sql,"encuestador", midefault)
	%>
	</TD>
</TR>
<TR>
	<TH>Sigla</TH>
	<TD><INPUT TYPE="text" NAME="sigla" id="sigla" value="<%=sigla%>" onclick="alert(Msemanas[0][0])"></TD>
</TR>
<TR>
	<TH>Rut&nbsp;Cliente</TH>
	<TD><INPUT TYPE="text" NAME="cliente" id="cliente" value="<%=cliente%>"></TD>
</TR>

<TR >
	<TH>Ciudad</TH>
	<TD>
		<% 
		sql="select nombre from sqlserver.desaerp.dbo.TMO_ciudades order by nombre"
		call milistbox(sql,"ciudad", miciudad)
		%>
	</TD>
</TR>

<TR>
	<TH>Cluster</TH>
	<TD>
	<% 
		midefault="Ninguno"
		sql="select nombre from sqlserver.desaerp.dbo.TMO_Clusters order by nombre"
		call milistbox(sql,"cluster", midefault)
	%>
	</TD>
</TR>
<TR>
	<TH>Canal</TH>
	<TD>
	<% 
		sql="select upper(nombre), nombre from sqlserver.desaerp.dbo.TMO_canales order by nombre"
		call milistbox(sql,"canal", midefault)
	%>
	</TD>
</TR>
<TR>
	<TH>Periodo</TH>
	<TD>
	<% 
		sql="select periodo from sqlserver.desaerp.dbo.tmo_semanas group by periodo"
		call milistbox(sql,"periodo", midefault)
	%>
	</TD>
</TR>
<TR>
	<TH>Semana</TH>
	<TD>
	<TABLE>
	<TR>
		<TD>	
		<% 
		sql="select 1 union select 2 union select 3 union select 4 union select 5"
		call milistbox(sql,"semana", midefault)
		%></TD>
		<TD><INPUT TYPE="text" NAME="desde" id="desde" size="10" readonly="true" style="text-align:center ; font-size:12px;background-color: #F4F4F4;border-width:0"></TD>
		<TD> a </TD>
		<TD><INPUT TYPE="text" NAME="hasta" id="hasta" size="10" readonly="true" style="text-align:center ; font-size:12px;background-color: #F4F4F4;border-width:0"></TD>
	</TR>
	</TABLE>
	</TD>
</TR>
</TABLE>
<HR>

<TABLE >
<%
	sql="SELECT TMO_PARAMETROSENC.idparametro, TMO_PARAMETROSENC.nombre, TMO_PARAMETROSENC.SeleccionMultiple, TMO_ENC_PREGUNTAS.Agrupa " & _
"FROM         sqlserver.desaerp.dbo.TMO_ENC_PREGUNTAS INNER JOIN " & _
"                      sqlserver.desaerp.dbo.TMO_PARAMETROSENC ON TMO_ENC_PREGUNTAS.idpregunta = TMO_PARAMETROSENC.idparametro " & _
"WHERE     (TMO_ENC_PREGUNTAS.idencuesta = 2) " & _
"ORDER BY TMO_ENC_PREGUNTAS.orden"

	grupoant = ""
	Set rs1=oConn.execute(sql)
	do until rs1.eof
		if len(rs1("agrupa")) = 0 then
			if grupoant<>rs1("agrupa") then
				%>
				<TR>
				<TD></TD>
				<TH><%=grupoant%></TH>
				<TD  border="1" style="border-collapse: collapse;border: 1px; border-style:solid; border-color: #003366" align="center">
				<TABLE  width="100%">
				<%
					a=split(txtgrupo,"|")
					'response.write()
					for x=0 to ubound(a)
						
						xb=split(a(x),"#")
						if ubound(xb) >0 then
							%><TR><TD><%=xb(0)%></TD><TD><%
							sql="select iditem, nombre, idclasificador from sqlserver.desaerp.dbo.TMO_PARAMETROSDET where idparametro='" & xb(1) & "' ORDER BY idparametro, idclasificador, nombre"
							if xb(2)=0 then
								call milistbox(sql   ,"cb_"  & right("000" & xb(1) ,3), midefault)
							else
								call checklistbox(sql,"clb_" & right("000" & xb(1) ,3), midefault)
							end if
						'vmulti=xb(2)
							%></TD></TR><%
						end if
						'response.write("<BR>" & ubound(xb))
					next
				%>
				</TABLE>
				</TD>
				</TR>
				<%
				txtgrupo=""
			end if
			%>
			<TR>
			<TD></TD>
			<TH><%=rs1("nombre")%></TH>
			<TD><%
				sql="select iditem, nombre, idclasificador from sqlserver.desaerp.dbo.TMO_PARAMETROSDET where idparametro='" & rs1("idparametro") & "' ORDER BY idparametro, idclasificador, nombre"
				if rs1("seleccionmultiple")=0 then
					call milistbox(sql   ,"cb_"  & right("000" & rs1("idparametro"),3), midefault)
				else
					call checklistbox(sql,"clb_" & right("000" & rs1("idparametro"),3), midefault)
				end if
			%></TD>
			</TR>
			<%
		else
			txtgrupo=txtgrupo & rs1("nombre") & "#" & rs1("idparametro") & "#" & rs1("seleccionmultiple") & "|"
		End if
		grupoant=rs1("agrupa")
		rs1.movenext
	loop

	'''''''''''''''''''''' valida si el ultimo es agrupado
	if len(txtgrupo)>0 then
	'response.write(txtgrupo)
				%>
				<TR>
				<TD></TD>
				<TH><%=grupoant%></TH>
				<TD  border="1" style="border-collapse: collapse;border: 1px; border-style:solid; border-color: #003366">
				<TABLE>
				<%
					a=split(txtgrupo,"|")
					'response.write()
					for x=0 to ubound(a)
						
						xb=split(a(x),"#")
						if ubound(xb) >0 then
							%><TR><TD><%=xb(0)%></TD><TD><%
							sql="select iditem, nombre, idclasificador from sqlserver.desaerp.dbo.TMO_PARAMETROSDET where idparametro='" & xb(1) & "' ORDER BY idparametro, idclasificador, nombre"
							if xb(2)=0 then
								call milistbox(sql   ,"cb_"  & right("000" & xb(1) ,3), midefault)
							else
								call checklistbox(sql,"clb_" & right("000" & xb(1) ,3), midefault)
							end if
						'vmulti=xb(2)
							%></TD></TR><%
						end if
						'response.write("<BR>" & ubound(xb))
					next
				%>
				</TABLE>
				</TD>
				</TR>
				<%
				txtgrupo=""
			end if
%>

</TABLE>

<BR>

<INPUT TYPE="submit" value='Guardar'>
</FORM>
<%
End sub 'encuesta()
'-----------------------------------------------------------------------------------------------------------------
sub guardar()
	
	encuestador = recuperavalor("encuestador")
	sigla       = recuperavalor("sigla"      )
	cluster     = recuperavalor("cluster"    )
	canal       = recuperavalor("canal"      )
	ciudad      = recuperavalor("ciudad"     )
	cliente     = recuperavalor("cliente"    )
	periodo     = recuperavalor("periodo"    )
	semana      = recuperavalor("semana"     )
	if len(trim(cliente))=0 then cliente=sigla
	
	idtipoencuesta="2"
	fecha = date()
	hora  = now()
	correlativo  = consultarapida("select isnull(max(idencuesta),0)+1 from sqlserver.desaerp.dbo.TMO_ENCUESTASENC")
	idencuestador= consultarapida("select idencuestador from sqlserver.desaerp.dbo.TMO_ENCUESTADORES where nombre='" & encuestador & "'")
	idcluster    = consultarapida("select idcluster from sqlserver.desaerp.dbo.tmo_clusters where nombre='" & trim(cluster) & "'")
	idcanal      = consultarapida("select idcanal   from sqlserver.desaerp.dbo.TMO_canales  where nombre='"  & trim(canal)   & "'")
	idciudad     = consultarapida("select idciudad  from sqlserver.desaerp.dbo.DIM_ciudades where nombre='" & trim(ciudad)  & "'")
	
	'response.write("<BR>" & correlativo )
	'response.write("<BR>" & sigla )
	'response.write("<BR>" & idencuestador )
	'response.write("<BR>" & idcluster )
	'response.write("<BR>" & fecha )
	'response.write("<BR>" & fecha )

	'Crea cliente
	sql="select top 1 * FROM sqlserver.desaerp.dbo.TMO_CLIENTES where idcliente='" & trim(cliente) & "' "
	Set rs=oConn.execute(Sql)
	if rs.eof then
		rs.Close
		rs.Open SQL, oConn, 1, 3
		rs.AddNew
		rs.fields("idcliente") = trim(cliente)
		rs.fields("sigla"    ) = sigla
		rs.fields("direccion") = ciudad
		rs.fields("idcanal"  ) = idcanal
		rs.fields("idciudad" ) = idciudad
		rs.fields("idcluster") = idcluster
		rs.Update
		rs.Close
	end if
	
	sql="select top 1 * FROM sqlserver.desaerp.dbo.TMO_ENCUESTASENC where id='" & trim(idtipoencuesta) & "' and idcliente='" & trim(sigla) & "' "
	Set rs=oConn.execute(Sql)
	if rs.eof then
		'response.write("idcanal : " & idcanal)
		rs.Close
		rs.Open SQL, oConn, 1, 3
		rs.AddNew
		rs.fields("idencuesta"   ) = correlativo
		rs.fields("id"           ) = idtipoencuesta
		rs.fields("idcliente"    ) = cliente
		rs.fields("sigla"        ) = sigla
		rs.fields("idencuestador") = idencuestador
		rs.fields("idcluster"    ) = idcluster
		rs.fields("idcanal"      ) = idcanal
		rs.fields("idciudad"     ) = idciudad
		rs.fields("fecha"        ) = fecha
		rs.fields("hora"         ) = hora
		rs.fields("idperiodo"    ) = periodo
		rs.fields("idsemana"     ) = semana

		'rs.fields("etiquetasvino") =etiquetasvino
		'rs.fields("etiquetasdesa") =etiquetasdesa
		rs.Update
		rs.Close
	else
		%><CENTER><BR><BR><BR>
		<FORM METHOD=POST ACTION="">
		<TABLE border="1" bgcolor="#FFFFFF">
		<TR bgcolor="#990000">
			<TD>Informe</TD>
		</TR>
		<TR>
			<TD><B>Ya hay una encuesta grabada para el cliente : <%=sigla%></B></TD>
		</TR>
		</TABLE><BR>
		<INPUT TYPE="submit" Value="Nueva Encuesta">
		</FORM>
		</CENTER><%
		exit sub
	end if
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
'------------------------------------------------------------------------------------
%>
</CENTER>
</BODY>
</HTML>


