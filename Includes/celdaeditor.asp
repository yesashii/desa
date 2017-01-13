<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<style type="text/css">
<!-- 
#nuevocomp{
	position:absolute;
}
#propimg{
	position:absolute;
}
#celdamatriz {
	background-color: #DCE1E9;
	border-style: solid;
	border-collapse: collapse;
}
#celdamatriz td{
	width: 24px;
	height: 30px;
	border-style: solid;
}
#celdamatriz tr{
	height: 24px;
}
body{
	font-family: verdana;
	font-size: 12px;
}
table{
	border-collapse: collapse ;
}
th{
	font-family: verdana;
}
td{
	font-family: verdana;
}
input {
	border: 1px #000000 solid;
}
#sombra {
	float:left;
	padding:0 5px 5px 0; /*Esta es la profundidad de nuestra sombra, sí haces más grandes estos valores, el efecto de sombra es mayor también */
	background: url(http://blogandweb.com/wp-content/uploads/2007/07/css-efecto-sombra.gif) no-repeat bottom right; /*Aquí es donde ponemos la imagen como fondo colocando su ubicación*/
} 

#idview {
	display:block;
	position:relative;
	top: -3px; /* Desfasamos la imagen hacia arriba */
	left:-3px; /*Desfasamos la imagen hacia la izquierda */
	padding:5px;
	background:#FFFFFF; /*Definimos un color de fondo */
	border:1px solid;
	border-color: #CCCCCC #666666 #666666 #CCCCCC /*Creamos un marco para acentuar el efecto */
}
	
-->
</style>
</HEAD>

<BODY topmargin="0" onmousemove="siguemouse()">
<CENTER>
<B><FONT SIZE="2" COLOR="#B6B8CB">Celda Editor</FONT></B>
<HR>
<TABLE>
<TR>
	<TD><% call celdamatriz() %></TD>
	<TD>Componentes</TD>
	<TD><% call menuherramientas() %></TD>
</TR>
</TABLE>

<TABLE>
<TR>
	<TD><BR><div id="sombra"><div id="idview"></div></div></TD>
	<TD>Lista Componentes</TD>
</TR>
</TABLE>

Comentario<BR>
<TEXTAREA NAME="" ROWS="5" COLS="50"></TEXTAREA>
<HR>
Finalezar

<Div id="propimg" style="FILTER: alpha(opacity=85)" style="display:none">
<TABLE border="1" style="border-color: #000000" bgcolor="#B9B9C8">
<TR>
	<TD colspan="4">
		<TABLE width="100%">
		<TR>
			<TH bgcolor="#000066" align="left"><button><B>[x]</B></button></TH>
			<TH colspan="3" bgcolor="#000066" align="left"><FONT SIZE="2" COLOR="#FFFFFF">Propiedades</FONT></TH>
		</TR>
		</TABLE>
	</TD>
</TR>
<TR>
	<TD>
		<TABLE>
		<TR>
			<TD><button onclick="fcortar()">cortar</button></TD>
			<TD><button>(-></button></TD>
			<TD><button><-)</button></TD>
		</TR>
		</TABLE>
	</TD>
	<TD>
		<TABLE>
		<TR>
			<TD colspan="2"><FONT SIZE="1" COLOR="#000033"><B>Nombre</B></FONT></TD>
		</TR>
		<TR>
			<TD><INPUT TYPE="text" NAME="" size="2"></TD>
			<TD></TD>
		</TR>
		</TABLE>
	</TD>
	<TD>
		<TABLE>
		<TR>
			<TD colspan="2"><FONT SIZE="1" COLOR="#000033"><B>Valor</B></FONT></TD>
		</TR>
		<TR>
			<TD><INPUT TYPE="text" id="mxy" size="2"></TD>
			<TD></TD>
		</TR>
		</TABLE>
	</TD>
	<TD>
		<button onclick="document.getElementById('propimg').style.display='none';preview()">Finalizar</button>
	</TD>
</TR>
</TABLE>
</Div>

<Div id="nuevocomp" style="FILTER: alpha(opacity=50)" style="display:none">
<IMG id="imgnuevocomp" SRC="" BORDER="1">
</Div>

</CENTER>
</BODY>
</HTML>
<%

sub celdamatriz()
%>
<div name="d_celdamatriz" id="d_celdamatriz" >
<TABLE name="celdamatriz" id="celdamatriz">
<%
for x1=1 to 12
	%><TR><%
	for x2=1 to 16
		%><TD id="td_<%=chr(x2+64) & x1%>"
			onmouseover="this.style.backgroundColor='#FFFFCC'" 
			onmouseout="this.style.backgroundColor='transparent'"
			onclick="guardaimg(this)"
		><IMG id="<%=chr(x2+64) & x1%>" SRC="" BORDER="0" style="display:none" ></TD><%
	next
	%></TR><%
next
%>
</TABLE>
</div>
<%
End sub
'------------------------------------------------------------------
sub menuherramientas()
%>

<TABLE>
<TR>
	<TD colspan="3"><FONT SIZE="1" COLOR="#000066"><B>Herramientas</B></FONT></TD>
</TR>
<TR>
	<TD>1</TD>
	<TD>2</TD>
	<TD>3</TD>
</TR>
<TR>
	<TD Colspan="3"></TD>
</TR>
<TR>
	<TD><IMG SRC="0001.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
	<TD><IMG SRC="0002.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
	<TD><IMG SRC="0003.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
</TR>
<TR>
	<TD><IMG SRC="0004.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
	<TD><IMG SRC="0005.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
	<TD><IMG SRC="0006.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
</TR>
<TR>
	<TD><IMG SRC="0007.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
	<TD><IMG SRC="0008.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
	<TD><IMG SRC="0009.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
</TR>
<TR>
	<TD><IMG SRC="0010.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
	<TD><IMG SRC="0011.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
	<TD><IMG SRC="0012.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
</TR>
<TR>
	<TD><IMG SRC="0013.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
	<TD><IMG SRC="0014.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
	<TD><IMG SRC="0015.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
</TR>
<TR>
	<TD><IMG SRC="0016.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
	<TD><IMG SRC="0017.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
	<TD><IMG SRC="0018.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
</TR>
<TR>
	<TD><IMG SRC="0019.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
	<TD><IMG SRC="0020.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
	<TD><IMG SRC="0021.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
</TR>
<TR>
	<TD><IMG SRC="0022.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
	<TD><IMG SRC="0023.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
	<TD><IMG SRC="0024.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
</TR>
<TR>
	<TD><IMG SRC="0025.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
	<TD><IMG SRC="0026.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
	<TD><IMG SRC="0027.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
</TR>
<TR>
	<TD><IMG SRC="0000.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
	<TD><IMG SRC="0000.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
	<TD><IMG SRC="0000.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
</TR>
<TR>
	<TD><IMG SRC="0000.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
	<TD><IMG SRC="0000.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
	<TD><IMG SRC="0000.BMP" BORDER="1" ALT="" onclick="activacomp(this)"></TD>
</TR>

</TABLE>
<%
end sub
'------------------------------------------------------------------

%>
<SCRIPT LANGUAGE="JavaScript">
<!--
var idimgactiva='';
//-----------------------------------------------------------------------------------
function preview(){
	txt='';
	//txt=document.getElementById('celdamatriz').rows[0].cells.length;
	xfilas=document.getElementById('celdamatriz').rows.length;
	xcolum=document.getElementById('celdamatriz').rows[0].cells.length;
	txt='<table border="0" cellpadding="0" cellspacing="0" style="border-collapse:collapse">'
	for (i=0;i<xfilas;i++){
		txt=txt+'<tr>';
		for (a=0;a<xcolum;a++){
			src_img='';
			id_img=String.fromCharCode((a+65))+(i+1);
			if (document.getElementById(id_img)){
				src_img=document.getElementById(id_img).src;
				if (src_img=='http://pda.desa.cl/includes/' || document.getElementById(id_img).style.display=='none'){
					src_img='http://pda.desa.cl/includes/9999.bmp';
				}
			}
			txt=txt+'<td><IMG SRC="'+src_img+'" WIDTH="12" HEIGHT="12" BORDER="0"></td>';
		}
		txt=txt+'</tr>';
	}
	txt=txt+'</table>';
	document.getElementById('idview').innerHTML=txt;
}
//-----------------------------------------------------------------------------------
function fcortar(){
	//alert(document.getElementById(idimgactiva).src);
	document.getElementById('imgnuevocomp').src=document.getElementById(idimgactiva).src
	document.getElementById('nuevocomp').style.display='';
	document.getElementById(idimgactiva).style.display='none';
	document.getElementById('propimg').style.display='none';
}
//-----------------------------------------------------------------------------------
function activacomp(miimagen){
	document.getElementById('imgnuevocomp').src=miimagen.src;
	document.getElementById('nuevocomp').style.display='';
	document.getElementById('propimg').style.display='none';	
}
//-----------------------------------------------------------------------------------
function siguemouse(){
	if(navigator.userAgent.indexOf("MSIE")>=0) {
		navegador=0;// IE
	}else{
		navegador=1;// Otros
	}

	if(navegador==0){
	 	cursorComienzoX=window.event.clientX+document.documentElement.scrollLeft+document.body.scrollLeft;
		cursorComienzoY=window.event.clientY+document.documentElement.scrollTop+document.body.scrollTop;
	}
	if(navegador==1){    
		cursorComienzoX=event.clientX+window.scrollX;
		cursorComienzoY=event.clientY+window.scrollY;
	}
	//cursorComienzoX=window.event.clientX+document.documentElement.scrollLeft+document.body.scrollLeft;
    //cursorComienzoY=window.event.clientY+document.documentElement.scrollTop+document.body.scrollTop;
	document.getElementById('nuevocomp').style.top=cursorComienzoY+10+'px';
	document.getElementById('nuevocomp').style.left=cursorComienzoX+10+'px';
	document.getElementById('mxy').value=cursorComienzoX;
//	alert(navegador);
//	alert(navigator.userAgent.indexOf("MSIE"));
}
//-----------------------------------------------------------------------------------
function guardaimg(miobjeto){
	miid=Right(miobjeto.id,Length(miobjeto.id)-3 );
	idimgactiva=miid;
	if (document.getElementById('nuevocomp').style.display == 'none'){
		if (document.getElementById(miid).style.display == 'none'){
			alert('Seleccione un Elemento');
		}else{
			document.getElementById('propimg').style.display='';
		}
	}else{
		document.getElementById(miid).src=document.getElementById('imgnuevocomp').src;
		document.getElementById(miid).style.display='';
		document.getElementById('propimg').style.top =document.getElementById('nuevocomp').style.top ;
		document.getElementById('propimg').style.left=document.getElementById('nuevocomp').style.left;
		document.getElementById('nuevocomp').style.display='none';
		document.getElementById('propimg').style.display='';
	}
	
}
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

//-->
</SCRIPT>