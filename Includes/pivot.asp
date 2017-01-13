<%
':simplepivot
' - Requiere Objeto de Conexion oConn

'----------------------------------------------------------------------------------------------------
function simplepivot(sql)
	'=====================================================================
	'|  simplepivot ver : 1.0 - 20 04 2009								|
	'=====================================================================
	' SQL 
	' Columna 0 : pivotes (columnas)
	' Columna 1 : filas pivot
	' Columna x : Datos columna 1
	' Ult colmn : Valores (datos)
'i=0

set frs=oConn.execute(sql)
if frs.eof then
	%><B>EOF: La Consulta SQL no Retorno ningun Resultado</B><%
else
	dim pvcolum()
	dim pvfilas()
	redim pvcolum(0)
	redim pvfilas(0)

	do until frs.eof
		'cargar vectores
		call agregadatovector(pvcolum, frs.fields(0))
		call agregadatovector(pvfilas, frs.fields(1))
	frs.movenext
	loop
'ordenar
	call arraySort(pvcolum,0)
	call arraySort(pvfilas,0)
'Construir pivot
	pvtxt="<TABLE border='1' style='border-collapse: collapse'><TR>"
	'TITULOS
	'Columnas fijas
	for x=1 to frs.Fields.count-2
		pvtxt=pvtxt & "<TD>" & frs.fields(x).name & "</TD>"
	next
	'Comumnas pivot
	for x=0 to ubound(pvcolum)
		'	response.write("<BR>" & pvcolum(x))
		pvtxt=pvtxt & "<TD style='writing-mode:tb-rl;'>" & pc(pvcolum(x),0) & "</TD>"
	next
	pvtxt=pvtxt & "</TR>"
	response.flush
	'CUERPO
	'Columnas fijas
	for x=0 to ubound(pvfilas)
		pvtxt=pvtxt & "<TR>"
		pvtxt=pvtxt & "<TD>" & pvfilas(x) & "</TD>"
		'call pvbuscadatocom(pvfilas(x))
'******
	frs.moveFirst
	do until frs.eof 'or frs.fields(1)=pvfilas(x)
		if frs.fields(1)=pvfilas(x) then
			'
			'response.write(frs.fields(2))
			exit do
		end if
	frs.movenext
	loop
	'frs.moveFirst
'******
		for pv_x=2 to frs.Fields.count-2 'frs.fields(pv_x)
			if frs.eof then
				pvtxt=pvtxt & "<TD> </TD>"
			else
				pvtxt=pvtxt & "<TD>" & pc(frs.fields(pv_x),0) & "</TD>"
			end if
		next
	'Comumnas pivot
		for pv_y=0 to ubound(pvcolum)'Datos
			'***************
			':: buscadato ::
			frs.moveFirst
			do until frs.eof '
				if ucase(frs.fields(1))=ucase(pvfilas(x)) then
					if ucase(frs.fields(0))=ucase(pvcolum(pv_y)) then '
						exit do
					end if
				end if
				'response.write("<BR>F:" & pvfilas(x) & " C:" & pvcolum(pv_y))
			frs.movenext
			loop
			'***************
			if frs.eof then
				pvtxt=pvtxt & "<TD>0</TD>"
			else
				pvtxt=pvtxt & "<TD>" & frs.fields(frs.Fields.count-1) & "</TD>"
			end if
		next
		pvtxt=pvtxt & "</TR>"
		response.write(pvtxt)
		pvtxt=""
		response.flush
	next
	pvtxt=pvtxt & "</TABLE>"

end if' : no es eof

response.write(pvtxt)
'control
	'for x=0 to ubound(pvcolum)
	'		response.write("<BR>(" & x & ") " & pvcolum(x))
	'next
	'for x=0 to ubound(pvfilas)
	'		response.write("<BR>(" & x & ") " & pvfilas(x))
	'next
End function 'simplepivot
'----------------------------------------------------------------------------------------------------
function agregadatovector(vectoraanalizar, vtvalor)
	existeenvector=0
	for x=0 to ubound(vectoraanalizar)
		if vtvalor=vectoraanalizar(x) then existeenvector=1
	next
	if existeenvector=0 then
		newx=ubound(vectoraanalizar)+1

		if ubound(vectoraanalizar)=0 then
			if len(vectoraanalizar(0))=0 then
				newx=0
			else
				Redim Preserve vectoraanalizar(newx)
			end if
		else
			Redim Preserve vectoraanalizar(newx)
		end if
		vectoraanalizar(newx)=vtvalor
	end if
End function
'----------------------------------------------------------------------------------------------------
function arraySort(vectorParaOrdenar,direccion) ' Le paso un 0 en el argumento Direccion para que el orden sea ascendente 
	max = uBound(vectorParaOrdenar) 
	For i=0 to max 
		For j = i+1 to max 
			if direccion = 0 then 'ordeno ascendente (menor a mayor) 
				if vectorParaOrdenar(i) > vectorParaOrdenar(j) then 
					Temp = vectorParaOrdenar(i) 
					vectorParaOrdenar(i) = vectorParaOrdenar(j) 
					vectorParaOrdenar(j) = Temp 
				end if 
			else 'ordeno descendente 
				if vectorParaOrdenar(i) < vectorParaOrdenar(j) then 
					Temp = vectorParaOrdenar(i) 
					vectorParaOrdenar(i) = vectorParaOrdenar(j) 
					vectorParaOrdenar(j) = Temp 
				end if 
			end if 
		next 
	next 
	arraySort = vectorParaOrdenar 
end function
'----------------------------------------------------------------------------------------------------
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
//Para todos los objetos fxlistbox
/*
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
*/
//-->
</SCRIPT>