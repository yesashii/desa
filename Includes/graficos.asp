<% 
'----------------------------------------------------------------------------------------------------------
function GraficoBarras(xtitulo,xsql,xwidth)

'Columna, valor1, Valor2 .....
'oConn.ConnectionTimeOut = 0
'oConn.CommandTimeout = 0
Set rs=oConn.execute(xsql)

masalto=0
do until rs.eof
	for xcolum=1 to rs.fields.Count - 1
		if cdbl(rs(xcolum))>masalto then masalto=cdbl(rs(xcolum))
		'response.write( "<BR>" & cdbl(rs(xcolum)) )
	next
rs.movenext
loop
'response.write( "<BR>mas alto : " & masalto )
xytitulo=replace(xtitulo," ","_")
%>
<div style="width:<%=xwidth%>px;background-color:#FFFFFF;" >

<CENTER>
<FONT SIZE="2" face="verdana" COLOR="#C0C0C0"><B><%=xtitulo%></B></FONT>
<TABLE border="1" width="100%" style='border-collapse: collapse; border-style: solid; border-color: #6699CC'>
<TR>
<%
rs.movefirst
do until rs.eof
	%><TH style='border-collapse: collapse; border-style: solid; border-color: #6699CC'><FONT SIZE="1" face="verdana" COLOR="#003366"><%=rs(0)%></FONT></TH><%
rs.movenext
loop
%>
</TR>
<TR>
<%
id_td=1
rs.movefirst
do until rs.eof
	%>
	<TD id="td_<%=id_td%>" valign="bottom" bgcolor="#E1F0FF" align="center" style='border-collapse: collapse; border-style: solid; border-color: #6699CC' 
	onMouseOver="
		this.style.backgroundColor='#E7FFE6'; 
		document.getElementById('resumen_<%=xytitulo%>').style.display='';
	var posx = 0;
    var posy = 0;
    if (!e) var e = window.event;
    if (e.pageX || e.pageY)
    {
        posx = e.pageX;
        posy = e.pageY;
    }
    else if (e.clientX || e.clientY)
    {
        posx = e.clientX;
        posy = e.clientY;
    }
		document.getElementById('resumen_<%=xytitulo%>').style.left = posx-10;
		document.getElementById('resumen_<%=xytitulo%>').style.top = posy+30;
		//transparent;
		//var elemDiv = document.getElementById('leyendaBDTCorte'); 
		//litCorte='valores falsos'; 
		//elemDiv.innerHTML+= litCorte;
		//alert(this.id)
		var id_td=Right(this.id,Length(this.id)-3)
		if (document.getElementById('c_<%=xytitulo%>_1')){
		document.getElementById('c_<%=xytitulo%>_1').innerHTML=recuperaalt(document.getElementById('img_<%=xytitulo%>'+id_td+'_1').alt,1);
		document.getElementById('v_<%=xytitulo%>_1').innerHTML=recuperaalt(document.getElementById('img_<%=xytitulo%>'+id_td+'_1').alt,2);
		}
		if (document.getElementById('c_<%=xytitulo%>_2')){
		document.getElementById('c_<%=xytitulo%>_2').innerHTML=recuperaalt(document.getElementById('img_<%=xytitulo%>'+id_td+'_2').alt,1);
		document.getElementById('v_<%=xytitulo%>_2').innerHTML=recuperaalt(document.getElementById('img_<%=xytitulo%>'+id_td+'_2').alt,2);
		}
		if (document.getElementById('c_<%=xytitulo%>_3')){
		document.getElementById('c_<%=xytitulo%>_3').innerHTML=recuperaalt(document.getElementById('img_<%=xytitulo%>'+id_td+'_3').alt,1);
		document.getElementById('v_<%=xytitulo%>_3').innerHTML=recuperaalt(document.getElementById('img_<%=xytitulo%>'+id_td+'_3').alt,2);
		}
		if (document.getElementById('c_<%=xytitulo%>_4')){
		document.getElementById('c_<%=xytitulo%>_4').innerHTML=recuperaalt(document.getElementById('img_<%=xytitulo%>'+id_td+'_4').alt,1);
		document.getElementById('v_<%=xytitulo%>_4').innerHTML=recuperaalt(document.getElementById('img_<%=xytitulo%>'+id_td+'_4').alt,2);
		}
		if (document.getElementById('c_<%=xytitulo%>_5')){
		document.getElementById('c_<%=xytitulo%>_5').innerHTML=recuperaalt(document.getElementById('img_<%=xytitulo%>'+id_td+'_5').alt,1);
		document.getElementById('v_<%=xytitulo%>_5').innerHTML=recuperaalt(document.getElementById('img_<%=xytitulo%>'+id_td+'_5').alt,2);
		}
		if (document.getElementById('c_<%=xytitulo%>_6')){
		document.getElementById('c_<%=xytitulo%>_6').innerHTML=recuperaalt(document.getElementById('img_<%=xytitulo%>'+id_td+'_6').alt,1);
		document.getElementById('v_<%=xytitulo%>_6').innerHTML=recuperaalt(document.getElementById('img_<%=xytitulo%>'+id_td+'_6').alt,2);
		}
		if (document.getElementById('c_<%=xytitulo%>_7')){
		document.getElementById('c_<%=xytitulo%>_7').innerHTML=recuperaalt(document.getElementById('img_<%=xytitulo%>'+id_td+'_7').alt,1);
		document.getElementById('v_<%=xytitulo%>_7').innerHTML=recuperaalt(document.getElementById('img_<%=xytitulo%>'+id_td+'_7').alt,2);
		}
		if (document.getElementById('c_<%=xytitulo%>_8')){
		document.getElementById('c_<%=xytitulo%>_8').innerHTML=recuperaalt(document.getElementById('img_<%=xytitulo%>'+id_td+'_8').alt,1);
		document.getElementById('v_<%=xytitulo%>_8').innerHTML=recuperaalt(document.getElementById('img_<%=xytitulo%>'+id_td+'_8').alt,2);
		}
		"
	onmouseout ="this.style.backgroundColor='#E1F0FF'; document.getElementById('resumen_<%=xytitulo%>').style.display='none'" >
		<TABLE border="0">
		<TR>
			<%
			for xcolum=1 to rs.fields.Count - 1 
			'response.write rs.fields.Count - 1 
				'H_Altura=altograf(cdbl(rs(xcolum)),masalto)
				src_color=filecolor(xcolum)
				if cdbl(rs(xcolum))=0 then src_color="http://pda.desa.cl/includes/blanco.bmp"

				%>
				<TD valign="bottom"><IMG id="img_<%=xytitulo%><%=id_td%>_<%=xcolum%>" SRC="<%=src_color%>" WIDTH="6" HEIGHT="<%=altograf(cdbl(rs(xcolum)),masalto)%>" BORDER="1" ALT="<%= rs(xcolum).name & " : " & formatnumber(rs(xcolum),0)%>"></TD>
				<%
				'response.write("<BR>" & altograf(cdbl(rs(xcolum)),masalto) )
			next
			%>
		</TR>
		</TABLE>
	</TD>
	<%
id_td=id_td+1
rs.movenext
loop
%>
</TR>
</TABLE>

<TABLE>
<TR>
	<%
	for xcolum=1 to rs.fields.Count - 1
		%>
		<TD>
			<TABLE>
			<TR>
				<TD><IMG SRC="<%=filecolor(xcolum)%>" WIDTH="6" HEIGHT="6" BORDER="1" ALT=""></TD>
				<TD><FONT SIZE="1" face="verdana" COLOR=""><B><%=rs.fields(xcolum).name%></B></FONT></TD>
			</TR>
			</TABLE>
		</TD>
		<%
	next
	%>
</TR>
</TABLE>
</CENTER>
	<div id="resumen_<%=xytitulo%>" name="resumen_<%=xytitulo%>" style="position:absolute; text-align:center; left:10px; top:100px; display:none;
		opacity: 0.8;
		filter: 'alpha(opacity=80)';
		filter: alpha(opacity=80);">
		<TABLE bgcolor="#E7FFE6" border="1" style='border-collapse: collapse; border-style: solid; border-color: #6699CC'>
			<%
			for xcolum=1 to rs.fields.Count - 1
			%>
			<TR>
				<TD><IMG SRC="<%=filecolor(xcolum)%>" WIDTH="6" HEIGHT="6" BORDER="1" ALT=""></TD>
				<TD><FONT SIZE="2" face="verdana" COLOR="#000066"><span id="c_<%=xytitulo%>_<%=xcolum%>"></span></FONT></TD>
				<TD><FONT SIZE="2" face="verdana" COLOR="#000066"><span id="v_<%=xytitulo%>_<%=xcolum%>"></span></FONT></TD>
			</TR>
			<%
			next
			%>
		
		</TABLE>
	</div>
</div>

<%
End function 'GraficoBarras
'----------------------------------------------------------------------------------------------------------
function GraficoBarrasCalendario(xtitulo,xsql,xwidth)
	'Fecha, valor1, Valor2 .....
%>
	<TABLE>
	<TR>
		<TD>Lunes</TD>
		<TD>Martes</TD>
		<TD>Miercoles</TD>
		<TD>Jueves</TD>
		<TD>Viernes</TD>
		<TD>Sabado</TD>
		<TD>Domingo</TD>
	</TR>
	<TR>
		<TD></TD>
		<TD></TD>
		<TD></TD>
		<TD></TD>
		<TD></TD>
		<TD>1</TD>
		<TD>2</TD>
	</TR>
	<TR>
		<TD>3</TD>
		<TD></TD>
		<TD></TD>
		<TD></TD>
		<TD></TD>
		<TD></TD>
		<TD></TD>
	</TR>
	<TR>
		<TD>10</TD>
		<TD></TD>
		<TD></TD>
		<TD></TD>
		<TD></TD>
		<TD></TD>
		<TD></TD>
	</TR>
	<TR>
		<TD>17</TD>
		<TD></TD>
		<TD></TD>
		<TD></TD>
		<TD></TD>
		<TD></TD>
		<TD></TD>
	</TR>
	<TR>
		<TD>24</TD>
		<TD></TD>
		<TD></TD>
		<TD></TD>
		<TD></TD>
		<TD></TD>
		<TD></TD>
	</TR>
	<TR>
		<TD>31</TD>
		<TD></TD>
		<TD></TD>
		<TD></TD>
		<TD></TD>
		<TD></TD>
		<TD></TD>
	</TR>
	</TABLE>
<%
End function 'GraficoBarrasCalendario
'----------------------------------------------------------------------------------------------------------
function altograf(xeval,xmax)
	on error resume next
	altograf=(xeval*100/xmax)*50/100
	'response.write("<BR>" & altograf)
	'altograf=1
end function
'----------------------------------------------------------------------------------------------------------
function filecolor(xindex)
	if xindex=0 then filecolor="http://pda.desa.cl/includes/amarillo.bmp"
	if xindex=1 then filecolor="http://pda.desa.cl/includes/azul.bmp"
	if xindex=2 then filecolor="http://pda.desa.cl/includes/rojo.bmp"
	if xindex=3 then filecolor="http://pda.desa.cl/includes/verde.bmp"
	if xindex=4 then filecolor="http://pda.desa.cl/includes/amarillo.bmp"
	if xindex=5 then filecolor="http://pda.desa.cl/includes/amarillo.bmp"
	if xindex=6 then filecolor="http://pda.desa.cl/includes/amarillo.bmp"
	if xindex=7 then filecolor="http://pda.desa.cl/includes/amarillo.bmp"
end function
'----------------------------------------------------------------------------------------------------------

%>
<SCRIPT LANGUAGE="JavaScript">
<!--
//-----------------------------------------------------------------------------------
function Length(str){return String(str).length;}
//-----------------------------------------------------------------------------------
function Left(str, n){
//window.alert(str + ':' + n);
	if (n <= 0)
	    return "";
	else if (n > String(str).Length)
	    return str;
	else
//		window.alert(3);
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
//-----------------------------------------------------------------------------------
function recuperaalt(str, n){
	if (n==1){
		vstr=Left(str,str.indexOf(':'));
	}else{
		vstr=Right(str,Length(str)-str.indexOf(':')-1);
	}
	return vstr;
}

//-->
</SCRIPT>