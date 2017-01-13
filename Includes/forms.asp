<%
':milistbox
':flabel
':fxlistbox
' - Requiere Objeto de Conexion oConn

'----------------------------------------------------------------------------------------------------
function milistbox(sql, nombre, default)
	'=====================================================================
	'|  milistbox ver : 1.0 - 05 06 2008								|
	'=====================================================================
	' SQL de 1 Columna , Valor=Col(1); Texto=Col(1)
	' SQL de 2 Columnas, Valor=Col(1); Texto=Col(2)
nombre=replace(trim(nombre)," ","_")
i=0

set frs=oConn.execute(sql)
if frs.eof then
	%><SELECT NAME="<%=nombre%>" id="<%=nombre%>" onChange="textbox_onChange('<%=nombre%>')">
			<option >EOF: La Consulta SQL no Retorno ningun Resultado</option>
	</SELECT><%
'exit function 
else
if frs.fields.count >1 then i=1
	%><SELECT NAME="<%=nombre%>" id="<%=nombre%>" onChange="textbox_onChange('<%=nombre%>')">
		<%
		encontrosel=0
		do until frs.eof 
			sel=""
			if lcase(frs.fields(0))=lcase(default) then 
				sel=" selected "
				encontrosel=1
			end if
			valor0=frs.fields(0)
			valor1=pc(frs.fields(i),0)
			if lcase(trim(valor0))="todos" then valor0=""
			%>
			<option <%=sel%>value="<%=valor0%>"><%=valor1%></option>
		<% frs.movenext 
		loop 
		if encontrosel=0 then
			'nada
		end if
		%>
	</SELECT><%
end if
End function 'milistbox
'----------------------------------------------------------------------------------------------------
function flabel(texto, Destino, Formato)
	'=====================================================================
	'| FLablel : Funcion constructora de Labels  | Ver 1.0 -  26/08/2008 |
	'=====================================================================
	'		Formato - Numerico = Tama�o de letra						|
	'               - Explicito : Font/style							|
	'		Destino = For												|
	'-------------------------------------------------------------------
	lformato=""
	if isnumeric(formato) then 
		ltamano=cint(formato)
		lformato=" face='verdana' SIZE='" & ltamano & "' COLOR='#000033' "
	else
		lformato=formato
	end if
	%><label id="lb_<%=Destino%>" name="lb_<%=Destino%>" for="<%=Destino%>">
	<FONT <%=lformato%>><%=pc(texto,0)%></FONT>
	</label><%
end function 'flabel
'----------------------------------------------------------------------------------------------------
Function fxlistbox(f_sql,f_nombre,f_default)
'f_sql="select 'Hola loco' as Saludo"
f_nombre=replace(trim(f_nombre)," ","_")
i=0

%>

<TABLE border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse'>
<TR>
	<TD><IMG SRC="/includes/Fundo_cb1.PNG" WIDTH="8" HEIGHT="18" BORDER="0" ALT=""></TD>
	<TD background="/includes/Fundo_cb2.PNG"><B><%=pc(f_nombre,1)%></B></TD>
	<TD><IMG SRC="/includes/Fundo_cb3.PNG" WIDTH="8" HEIGHT="18" BORDER="0" ALT=""></TD>
</TR>
<TR>
	<TD background="/includes/Fundo_cb4.PNG"></TD>
	<TD background="/includes/Fundo_cb5.PNG">
<% 'call milistbox(f_sql, f_nombre, f_default) 
	
set frs=oConn.execute(f_sql)
if frs.eof then
	%>EOF: La Consulta SQL no Retorno ningun Resultado<%
else
	if frs.fields.count >1 then i=1
	%><SELECT NAME="S_<%=f_nombre%>" id="<%=f_nombre%>" onChange="fjs_combochange(this.id)" style="background-color: #D9ECFF">
		<option value="todos">Todos</option>
		<option value="Selec">Seleccion Multiple</option><%
		do until frs.eof 
			sel=""
			if lcase(frs.fields(0))=lcase(default) then sel=" selected "
			valor0=frs.fields(0)
			valor1=pc(frs.fields(i),0)
			if lcase(trim(valor0))="todos" then valor0=""
			%><option <%=sel%>value="<%=valor0%>"><%=valor1%></option><%
			frs.movenext 
		loop 
	%></SELECT><%
end if
	
	%></TD>
	<TD background="/includes/Fundo_cb6.PNG"></TD>
</TR>
<TR>
	<TD background="/includes/Fundo_cb4.PNG"></TD>
	<TD background="/includes/Fundo_cb11.PNG">
	<TABLE id="l_<%=f_nombre%>" width="100%" border='1' cellpadding='0' cellspacing='0' style='border-collapse: collapse' style="cursor:pointer; display:none;">
	<%
	'style="cursor:pointer; display:none;"
		frs.movefirst
		fx=0
		do until frs.eof 
			sel=""
			if lcase(frs.fields(0))=lcase(default) then sel=" selected "
			valor0=frs.fields(0)
			valor1=pc(frs.fields(i),0)
			if lcase(trim(valor0))="todos" then valor0=""
			%><TR onMouseOver="this.style.backgroundColor='#D9ECFF'" 
				  onmouseout ="this.style.backgroundColor='transparent'" ><TD>
			<INPUT TYPE="checkbox" NAME="<%=f_nombre%>" ID="<%=f_nombre & fx %>" value="<%=valor0%>" checked><%
			call flabel(valor1, f_nombre & fx, 2)
			%>&nbsp;</TD></TR><%
			fx=fx+1
			frs.movenext 
		loop 
	%>
	</TABLE>
	</TD>
	<TD background="/includes/Fundo_cb6.PNG"></TD>
</TR>
<TR>
	<TD><IMG SRC="/includes/Fundo_cb7.PNG" WIDTH="8" HEIGHT="10" BORDER="0" ALT=""></TD>
	<TD background="/includes/Fundo_cb8.PNG"></TD>
	<TD><IMG SRC="/includes/Fundo_cb9.PNG" WIDTH="8" HEIGHT="10" BORDER="0" ALT=""></TD>
</TR>
</TABLE><%
End Function 'fxlistbox
'----------------------------------------------------------------------------------------------------
Function checklistbox(f_sql,f_nombre,f_default)
	set frs=oConn.execute(f_sql)
	if frs.fields.count >2 then 'i=1
		xfamiliaant="nada"
		'xfamiliaant=frs(2)
		txttitulo=""
		txtdetalle=""
		do until frs.eof
			if xfamiliaant<>frs(2) then
				txttitulo=txttitulo & "<TH bgcolor='#EFEFFC'>" & frs(2) & "</TH>"
				if xfamiliaant<>"nada" then txtdetalletotal=txtdetalletotal & "<TD valign='top'>" & txtdetalle & "</TD>"
				txtdetalle=""
			end if
			txtdetalle=txtdetalle & "<INPUT TYPE='checkbox' NAME='" & f_nombre & right("000" & frs(0),3) & "' ID='" & f_nombre & "_" & frs(0) & "'><label for='" & f_nombre & "_" & frs(0) & "' >"& replace(frs(1)," ","&nbsp;") &"</label><BR>"
			xfamiliaant=frs(2)
		frs.movenext
		loop
				if xfamiliaant<>"nada" then txtdetalletotal=txtdetalletotal & "<TD valign='top'>" & txtdetalle & "</TD>"
		%>
		<TABLE border="1" style="border-collapse: collapse;border: 1px; border-style:solid; border-color: #003366">
<!-- 		<TR>
		<%=txttitulo%>
		</TR> -->
		<TR>
		<%=txtdetalletotal%>
		</TR>
		</TABLE>
		<%
	end if

	%>
<!-- 	 <label for="pallets">nombre</label>  -->
	<%
	
End Function 'checklistbox
'----------------------------------------------------------------------------------------------------
Function checklistbox2(f_sql,f_nombre,f_default)
	set frs=oConn.execute(f_sql)
	if frs.fields.count >2 then 'i=1
		%>
<TABLE border="1" style="border-collapse: collapse;border: 1px; border-style:solid; border-color: #003366">
<TR>
	<TD>
	<DIV id="Layer1" style= "width:180px; height: 100px; overflow-y: scroll ; overflow-x: none; ">
	<TABLE width="100%">
	<%
		do until frs.eof
		%><TR bgcolor='' 
		onMouseOver="this.style.backgroundColor='#EFEFEF'" 
		onmouseout="this.style.backgroundColor=''" >
			<TD><INPUT TYPE='checkbox' NAME='<%=f_nombre & right("000" & frs(0),3)%>' ID="<%=f_nombre & "_" & frs(0) %>"></TD>
			<TD><label for="<%=f_nombre & "_" & frs(0) %>"><%=replace(frs(1)," ","&nbsp;")%></label></TD>
			<TD><%=replace(frs(2)," ","&nbsp;")%></TD>
		</TR><%
		frs.movenext
		loop
	%>
	</TABLE>
	</DIV>
	</TD>
</TR>
</TABLE>
		<%
	end if
	
End Function 'checklistbox2
'----------------------------------------------------------------------------------------------------
Function checklistbox3(f_sql,f_nombre,f_default)
'Seleccion multible con indice
	set frs=oConn.execute(f_sql)
	if frs.fields.count >2 then 'i=1
		%>


<TABLE border="1" style="border-collapse: collapse;border: 1px; border-style:solid; border-color: #003366">
<TR>
	<TD>

	<!-- <DIV id="Layer2" style= "width:40px; height: 100px; overflow-y: scroll ; overflow-x: none; "> -->
	<TABLE width="100%" border="1" style="border-collapse: collapse;border: 1px; border-style:solid; border-color: #003366">
	<TR><TD align='center' bgcolor='' 
		onMouseOver="this.style.backgroundColor='#D1DAE9'" 
		onmouseout="this.style.backgroundColor=''" 
		onClick="fjs_filtrarchecklist('<%=f_nombre%>','todos')"><B>*</B></TD>
	<%
	indice_anterior="nadaescomoesto"
	do until frs.eof
		if indice_anterior<>frs(2) then
		%><TD align='center' bgcolor='' 
		onMouseOver="this.style.backgroundColor='#D1DAE9'" 
		onmouseout="this.style.backgroundColor=''"
		onClick="fjs_filtrarchecklist('<%=f_nombre%>','<%=frs(2)%>')"
		><B><%=replace(frs(2)," ","&nbsp;")%></B></TD><%
		end if
		indice_anterior=frs(2)
		frs.movenext
	loop
	%>
	</TR>
	</TABLE>
	<!-- </DIV> -->

	</TD>
</TR>
<TR>
<TD>
	<DIV id="Layer1" style= "width:180px; height: 100px; overflow-y: scroll ; overflow-x: none; ">
	<TABLE width="100%" ><TR><TD>
	<%
		set frs=oConn.execute(f_sql)
		indice_anterior="nadaescomoesto"
		do until frs.eof
		'Crea DIV
			if indice_anterior<>frs(2) then
				if indice_anterior="nadaescomoesto" then
					%><DIV name="DIV_<%=f_nombre%>" id="DIV_<%=f_nombre%>_<%=frs(2)%>"><TABLE width="100%" ><%
				else
					%></TABLE></DIV><DIV name="DIV_<%=f_nombre%>" id="DIV_<%=f_nombre%>_<%=frs(2)%>"><TABLE width="100%" ><%
				end if
			end if
		'linea
		%><TR 
		bgcolor='' 
		onMouseOver="this.style.backgroundColor='#EFEFEF'" 
		onmouseout="this.style.backgroundColor=''" >
			<TD><INPUT TYPE='checkbox' NAME='<%=f_nombre & right("000" & frs(0),3)%>' ID="<%=f_nombre & "_" & frs(0)  %>"></TD>
			<TD><label for="<%=f_nombre & "_" & frs(0) %>"><%=replace(frs(1)," ","&nbsp;")%></label></TD>
			<TD><%=replace(frs(2)," ","&nbsp;")%></TD>
		</TR><%

		indice_anterior=frs(2)
		frs.movenext
		loop
		%></TABLE></DIV><%
	%>
		
	</TD></TR></TABLE>






	</DIV>
	</TD>
</TR>
</TABLE>
		<%
	end if
	
End Function 'checklistbox3
'----------------------------------------------------------------------------------------------------
Function checklistbox4(f_sql,f_nombre,f_default)
'Seleccion multible con indice
	set frs=oConn.execute(f_sql)
	if frs.fields.count >2 then 'i=1
		%>


<TABLE border="1" style="border-collapse: collapse;border: 1px; border-style:solid; border-color: #003366">
<TR>
	<TD>

	<!-- <DIV id="Layer2" style= "width:40px; height: 100px; overflow-y: scroll ; overflow-x: none; "> -->
	<TABLE width="100%" border="1" style="border-collapse: collapse;border: 1px; border-style:solid; border-color: #003366">
	<TR><TD align='center' bgcolor='' 
		onMouseOver="this.style.backgroundColor='#D1DAE9'" 
		onmouseout="this.style.backgroundColor=''" 
		onClick="fjs_filtrarchecklist('<%=f_nombre%>','todos')"><B>*</B></TD>
	<%
	indice_anterior="nadaescomoesto"
	do until frs.eof
		if indice_anterior<>frs(2) then
		%><TD align='center' bgcolor='' 
		onMouseOver="this.style.backgroundColor='#D1DAE9'" 
		onmouseout="this.style.backgroundColor=''"
		onClick="fjs_filtrarchecklist('<%=f_nombre%>','<%=frs(2)%>')"
		><B><%=replace(frs(2)," ","&nbsp;")%></B></TD><%
		end if
		indice_anterior=frs(2)
		frs.movenext
	loop
	%>
	</TR>
	</TABLE>
	<!-- </DIV> -->

	</TD>
</TR>
<TR>
<TD>
	<DIV id="Layer1" style= "width:180px; height: 100px; overflow-y: scroll ; overflow-x: none; ">
	<TABLE width="100%" ><TR><TD>
	<%
		set frs=oConn.execute(f_sql)
		indice_anterior="nadaescomoesto"
		do until frs.eof
		'Crea DIV
			if indice_anterior<>frs(2) then
				if indice_anterior="nadaescomoesto" then
					%><DIV name="DIV_<%=f_nombre%>" id="DIV_<%=f_nombre%>_<%=frs(2)%>"><TABLE width="100%" ><%
				else
					%></TABLE></DIV><DIV name="DIV_<%=f_nombre%>" id="DIV_<%=f_nombre%>_<%=frs(2)%>"><TABLE width="100%" ><%
				end if
			end if
		'linea
		%><TR 
		bgcolor='' 
		onMouseOver="this.style.backgroundColor='#EFEFEF'" 
		onmouseout="this.style.backgroundColor=''" >
			<TD><INPUT TYPE='checkbox' NAME='<%=f_nombre & right("000" & frs(0),3)%>' ID="<%=f_nombre & "_" & frs(0)  %>"></TD>
			<TD><label for="<%=f_nombre & "_" & frs(0) %>"><%=replace(frs(1)," ","&nbsp;")%></label></TD>
			<TD><%=replace(frs(2)," ","&nbsp;")%></TD>
		</TR><%

		indice_anterior=frs(2)
		frs.movenext
		loop
		%></TABLE></DIV><%
	%>
		
	</TD></TR></TABLE>






	</DIV>
	</TD>
</TR>
</TABLE>
		<%
	end if
	
End Function 
'----------------------------------------------------------------------------------------------------

%>
<SCRIPT LANGUAGE="JavaScript">
<!--
//Para todos los objetos fxlistbox
function fjs_combochange(mi_id){
	var combo=document.getElementById(mi_id).value;
	//des seleccionar todo ???
		for (var i=document.getElementById(mi_id).options.length-1; i>=0; i--){
			if (document.getElementById(mi_id+(i-2)) ){
				document.getElementById(mi_id+(i-2)).checked=false;
			}
		}
		//alert();
	if(combo=='Selec'){
		document.getElementById('l_'+mi_id).style.display='';
	}else{
		document.getElementById('l_'+mi_id).style.display='none';
	}
	if(combo=='todos'){
		//selecciona todos
		
		for (var i=document.getElementById(mi_id).options.length-1; i>=0; i--){
			if (document.getElementById(mi_id+(i-2)) ){
				document.getElementById(mi_id+(i-2)).checked=true;
			}
		}
	}else{
		//Selecciona = combo
		  for (var i=document.getElementById(mi_id).options.length-1; i>=0; i--){
			if (combo==document.getElementById(mi_id).options[i].value){//si lo encuentra
				if (document.getElementById(mi_id+(i-2)) ){//Si Existe
					document.getElementById(mi_id+(i-2)).checked=true;
				}
			}
		}
	}	
//alert(combo);
}
//------------------------------------------------------------------------------------------
function fjs_filtrarchecklist(Ncheck, filtro){
	//alert(Ncheck);
	//alert(filtro);
	//var mi_id='DIV_'+Ncheck+'_'+filtro;
	//document.getElementById(mi_id).style.display='';
 
	capas=document.getElementsByTagName('div');
	//alert(capas.length);
	for (i=0;i<capas.length;i++){
		if(capas[i].id.indexOf('DIV_'+Ncheck) != -1){
			capas[i].style.display='none';
		}
	}
	if(document.getElementById('DIV_'+Ncheck+'_'+filtro)){
		document.getElementById('DIV_'+Ncheck+'_'+filtro).style.display='';
	}else{
		//alert("la capa no existe");
	}

	if (filtro=='todos'){
		//alert(filtro);
		for (i=0;i<capas.length;i++){
			if(capas[i].id.indexOf('DIV_'+Ncheck) != -1){
				capas[i].style.display='';
			}
		}
	}

}
//-->
</SCRIPT>