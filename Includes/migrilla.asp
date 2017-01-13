<SCRIPT LANGUAGE="JavaScript">
var str='';
var colorcuerpo='';
//var cols=MatrizTL.length;
function Right(str0, n){
if (n<=0) return "";
else if (n>String(str0).length)return str0;
else {var iLen=String(str0).length;
return String(str0).substring(iLen,iLen-n);}
}
/*
TramaCabeza
ColorCabeza
ColorCuerp0
ColorCuerp1
ColorCuerp2
*/
//--------------------------------------------------------------------------------------------------------
function tablaJava0(){
	str='';
	str=str+'<TABLE '+formato+'>';
	//Encabezado
	  str=str+'<TR bgcolor='+ColorCabeza+'>';
	   for(i=0;i<MatrizTL.length;i++){
		str=str+'<TD align="center" background='+TramaCabeza+'><B>&nbsp;'+ MatrizTL[i] +'&nbsp;</B></TD>';
	   };
	str=str+'</TR>';
	//Cuerpo
	for(i=0;i<=200;i++){
	colorcuerpo=ColorCuerp1;
	u=Right(i,1);
	if(u==2|u==4|u==6|u==8|u==0){colorcuerpo=ColorCuerp2};
	  str=str+'<TR bgcolor='+colorcuerpo+' onMouseOver="this.style.backgroundColor=\''+ColorCuerp0+'\'" onmouseout="this.style.backgroundColor=\''+colorcuerpo+'\'" >';
	  for(e=0;e<MatrizTL.length;e++){
	    str=str+'<TD>&nbsp;'+ Matriz2D[i][e] +'&nbsp;</TD>';
	  }
	  str=str+'</TR>';
	  if(Matriz2D[i+1][0]==undefined){i=200};
	};
	str=str+'</TABLE>';
	document.getElementById("tbljava").innerHTML=str
}//function tablaJava0()
//--------------------------------------------------------------------------------------------------------
function tablaJava(){
	str='';
	str=str+'<TABLE '+formato+'>';
	//Encabezado
	  str=str+'<TR bgcolor='+ColorCabeza+'>';
	   for(i=0;i<MatrizTL.length;i++){
		str=str+'<TD align="center" background='+TramaCabeza+'><B>&nbsp;'+ MatrizTL[i] +'&nbsp;</B></TD>';
	   };
	  str=str+'</TR>';
	//Filtro
	  str=str+'<TR bgcolor='+ColorCabeza+'>';
	   for(i=0;i<MatrizTL.length;i++){
		str=str+'<TD align="center" background='+TramaCabeza+'><INPUT TYPE="text" size="4" NAME="" style="background-color: #F3F3F3 ; border: 1 solid '+ColorCabeza+'"><select size="1" name="D1"></select></TD>';
	   };
	str=str+'</TR>';
	//Cuerpo
	for(i=0;i<=200;i++){
	colorcuerpo=ColorCuerp1;
	u=Right(i,1);
	if(u==2|u==4|u==6|u==8|u==0){colorcuerpo=ColorCuerp2};
	  str=str+'<TR bgcolor='+colorcuerpo+' onMouseOver="this.style.backgroundColor=\''+ColorCuerp0+'\'" onmouseout="this.style.backgroundColor=\''+colorcuerpo+'\'" >';
	  for(e=0;e<MatrizTL.length;e++){
	    str=str+'<TD>&nbsp;'+ Matriz2D[i][e] +'&nbsp;</TD>';
	  }
	  str=str+'</TR>';
	  if(Matriz2D[i+1][0]==undefined){i=200};
	};
	str=str+'</TABLE>';
	document.getElementById("tbljava").innerHTML=str
}//function tablaJava()
//--------------------------------------------------------------------------------------------------------
</SCRIPT>
<% on error resume next
':::::::::::::::::: conexion :::::::::::::::::
public oConn, rs
'Public SQL, Estilo, Modo, formato
Public TramaCabeza, ColorCabeza, ColorCuerp0, ColorCuerp1, ColorCuerp2, ColorTxt0, elink, eEspecial
Set oConn = server.createobject("ADODB.Connection")
oConn.ConnectionTimeOut = 0
oConn.CommandTimeout = 0
oConn.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=todo;User Id=sa;Password=desakey;"

'--------------------------------------------------------------------------------------------
Sub Migrilla(SQL, Estilo, Modo, formato, Especial)
'|---------------------------------------------------------------------|
'| Ver 1.0 | 29 Ene 2008 | Simon Hernandez | robotekmania@hotmail.com  |
'|---------------------------------------------------------------------|
'| SQL     : String (Sentencia transac) ok                             |
'| Estilo  : 0 = tablas Simple          ok                             |
'|           1 = Escala Grises          ok                             |
'|           2 = Selestes               ok                             |
'|           3 = Verdes                 ok                             |
'| Modo    : 0 = Nada                                                  |
'|           1 = Filtro Encabezados                                    |
'|           2 = Tabla Con Editor                                      |
'|           3 = Listado (aprupa por)                                  |
'|--------------(Configuracion Especial por nombre de Campo)-----------|
'|Columnas :CK = checkbox Dinamico (Campo Id Obligatorio)    ok (falta Java)       |
'|          OP = radio Dinamico (Campo Id Obligatorio)       ok (Falta Java)
'|        LinK =
'|--------------------------------------------------------------------|

eLink=Especial
eEspecial=Especial
if lcase(Especial)="onclick" then eLink=""
if lcase(Especial)="formatnumber" then eLink=""

'response.write(Especial)
':validando
	if len(trim(sql))=0      then exit sub
	on error resume next
	set rs=oConn.execute(Sql)
	if err.number <>0 then
		response.write("<BR>Error Numero = " & Err.number)
		response.write("<BR>Descripcion. = " & Err.Description)
		response.write("<BR>Fuente = " & Err.Source)
		response.write("<BR><BR><B>El servidor No Responde, Intente En Otro Momento ...</B>")
	end if
	if rs.eof or rs.bof      then exit sub
	if not isnumeric(Estilo) then exit sub
	if not isnumeric(Modo)   then exit sub
	if len(trim(formato))= 0 then formato="border='1' cellpadding='0' cellspacing='0' style='border-collapse: collapse; border-left-width: 0; border-top-width: 0' bordercolor='#808080'"
	elink=" onclick='window.open("  & chr(34) & elink & chr(34) & ")'"
':Estilo
	Select Case(Estilo)
	Case 1:
		TramaCabeza="/Includes/tazul.bmp" : ColorCabeza="#C0C0C0" : ColorTxt0="#000033"
		ColorCuerp0="#CCCCCC"   : ColorCuerp1="#DDDDDD" : ColorCuerp2="#FFFFFF"
	Case 2:
		TramaCabeza="/Includes/tazul.bmp" : ColorCabeza="#C4D3F2" : ColorTxt0="#000033"
		ColorCuerp0="#FFFFFF"   : ColorCuerp1="#ECEEFF" : ColorCuerp2="#E8EDF4"
	Case 3:
		TramaCabeza="/Includes/tazul.bmp" : ColorCabeza="#CCFF99" : ColorTxt0="#000033"
		ColorCuerp0="#CACACA"   : ColorCuerp1="#E6FFE6" : ColorCuerp2="#E8EDF4"
	Case 4:
		TramaCabeza="" : ColorCabeza="#000066" : ColorTxt0="#FFFFFF"
		ColorCuerp0=""   : ColorCuerp1="" : ColorCuerp2=""
	End Select

	response.write("<TABLE border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse;'><TR><TD>" & chr(10))'borde tabla ini
	response.flush()
':Modo
	Select Case(Modo)
	Case 0:
		Call TablaHtml(formato)
	Case 1:
		response.write("<span id='tbljava'></span>" & chr(10))
		i=1000
		response.write("<SCRIPT LANGUAGE=" & chr(34) & "JavaScript" & chr(34) &  "> " & chr(10) )	
		' " TYPE=" & chr(34) & "text/javacript" & chr(34) & 
		response.write("var MatrizTL=new Array(" & rs.fields.count & ");" & chr(10) )
		  for x=0 to rs.fields.count-1
			response.write("MatrizTL[" & x & "]='" & replace(pc(rs.fields(x).name,2),"'",chr(34)) & "';" & chr(10) )
		  next
		response.write("var Matriz2D=new Array(" & i & ");" & chr(10) )
		response.write("for(i=0;i<=" & i & ";i++){Matriz2D[i]=new Array(7);}" & chr(10) )
		rs.movefirst
		a=0
		do until rs.eof
		  for x=0 to rs.fields.count-1
			response.write("Matriz2D[" & a & "][" & x & "]='" & replace(pc(rs.fields(x),2),"'",chr(34)) & "';" & chr(10) )
		  next
		rs.movenext
		a=a+1
		loop
		
		response.write("var TramaCabeza='" & TramaCabeza & "';" & chr(10) )
		response.write("var ColorCabeza='" & ColorCabeza & "';" & chr(10) )
		response.write("var ColorCuerp0='" & ColorCuerp0 & "';" & chr(10) )
		response.write("var ColorCuerp1='" & ColorCuerp1 & "';" & chr(10) )
		response.write("var ColorCuerp2='" & ColorCuerp2 & "';" & chr(10) )
		
		response.write("var formato='" & replace(formato,"'",chr(34) ) & "';" & chr(10) )
		response.write("tablaJava();" & chr(10) )
		response.write("</SCRIPT>" & chr(10) )
		
	End Select
	
  response.write("</TD><TD background='/Includes/derecha.gif'>&nbsp;</TD></TR><TR><TD background='/Includes/bajo.gif'>&nbsp;</TD><TD background='/Includes/Esquina.gif'>&nbsp;</TD></TR></TABLE>" & chr(10) & chr(10))'borde tabla fin

'response.write("<TABLE><TR><TD onclick='tablaJava()'>tabla</TD></TR></TABLE>")
rs.close
End sub 'Migrilla
'--------------------------------------------------------------------------------------------
Sub TablaHtml(formato)
	on error resume next
	response.write( "<TABLE " & formato & ">" & chr(10))
	':::::Titulos
	response.write("<TR>")
	For x=0 to rs.Fields.count-1
		NombreCol=rs.Fields(x).Name
		%><TD align='center' background="<%=TramaCabeza%>" bgcolor='<%=ColorCabeza%>' ><B>&nbsp;<%=replace(PC(NombreCol,2),"#000033",ColorTxt0)%>&nbsp;</B></TD><%
		response.write(chr(13))
	next
	response.write("</TR>")
	response.flush()
	':::::Datos
	x1=0
	if len(elink)=26 then elink=""
	rlink=elink
	Do until rs.eof 
		elink=rlink
		'elink=replace(elink,"rs.fields(0)",rs.fields(0))
		'if rs.fields.count>1 then elink=replace(elink,"rs.fields(1)",rs.fields(1))
		'if rs.fields.count>1 then elink=replace(elink,"rs.fields(2)",rs.fields(2))
		'if rs.fields.count>1 then elink=replace(elink,"rs.fields(3)",rs.fields(3))
		
		for xl=0 to rs.fields.count-1
			elink=replace(elink,"rs.fields(" & xl & ")",rs.fields(xl))
		next
		elink=replace(elink,".asp?",".asp?time=" & time() & "&")
		if lcase(eEspecial)="onclick" then elink=" onclick='migrilla_onclick(this)'"
		'eLink="onclick=migrilla_onclick(this)"
		'elink=len(eEspecial)
		Colorcuerpo=ColorCuerp1
		if x1 mod 2 =0 then Colorcuerpo=ColorCuerp2
		%><TR bgcolor='<%=ColorCuerpo%>' 
		onMouseOver="this.style.backgroundColor='<%=Colorcuerp0%>'" 
		onmouseout="this.style.backgroundColor='<%=Colorcuerpo%>'" 
		<%=elink%>
		>
		<%
		For x=0 to rs.Fields.count-1
			DatoCol=""
			if not isnull(rs.Fields(x).value) then DatoCol=trim(rs.Fields(x).value)
			'if len(DatoCol)=0 then DatoCol=" "
			'DatoCol="1"
			Mialign="left"
			if isnumeric(Datocol) then 
				Mialign="right"
				if eEspecial="formatnumber" then Datocol=formatnumber(Datocol,0)
			end if

			DatoCol=PC(DatoCol,2)
			if not isnull(rs.Fields(x).value) then mval=replace(rs.Fields(x).value," ","_")
			if ucase(rs.Fields(x).name)="CK" then DatoCol="<input type='checkbox' name='" & mval & "' Id='" & mval & "'>"
			if ucase(rs.Fields(x).name)="OP" then DatoCol="<input type='radio' value='" & mval & "' name='op' id='op'>"
			response.write("<TD align='" & Mialign & "' style='padding-left:6px; padding-right:5px;' >" & DatoCol & "</TD>" & chr(13) )'Esta es la que se escribe
		next
		response.write("</TR>" & chr(13) )
	x1=x1+1
	rs.movenext
	loop
	response.write("</TABLE>" & chr(10))
End Sub 'TablaHtml()
'--------------------------------------------------------------------------------------------
Function PC(sString, mzise) 
  Dim sWhiteSpace, bCap, iCharPos, sChar 
  sWhiteSpace = Chr(32) & Chr(9) & Chr(13) 
  sString = LCase(sString)
  bCap = True 
  For iCharPos = 1 to Len(sString) 
    sChar = Mid(sString, iCharPos, 1) 
    If bCap = True Then 
      sChar = UCase(sChar) 
    End If 
    ProperCase = ProperCase + sChar 
    If InStr(sWhiteSpace, sChar) Then 
      bCap = True 
    Else 
      bCap = False 
    End If 
  Next  
PC="<FONT SIZE='" & Mzise & "' face='arial' COLOR='#000033'>" & replace(ProperCase," ","&nbsp;") & "</FONT>"
PC=replace(replace(PC,"%20"," "),"***",chr(34))
if mzise=0 then PC=ProperCase
End Function 
'--------------------------------------------------------------------------------------------
Function cisnull(valor, reemplazo)
	if isnull(valor) then
		cisnull=reemplazo
	else
		cisnull=valor
	end if
End Function
'--------------------------------------------------------------------------------------------
Function iif(consulta, valortrue, valorfalse)
	if consulta then
		iif=valortrue
	else
		iff=valorfalse
	end if
End Function
'--------------------------------------------------------------------------------------------
Function recuperavalor(objeto)
	if len(request.querystring(objeto))=0 then
		if len(request.form(objeto))>0 then
			recuperavalor=request.form(objeto)
		else
			recuperavalor=""
		end if
	else
		recuperavalor=request.querystring(objeto)
	end if
End Function
'--------------------------------------------------------------------------------------------
%>