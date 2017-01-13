<%
dim vend, cliente

vend    = request.querystring("vend")
vend=replace(vend,"NN","Ñ")'Error blackberry
vend=replace(vend,"?","Ñ")'Error blackberry
cliente = trim(request.QueryString("cliente"))

%>
<html>
<head>
<title>noventa</title>
</head>
<body text="#000000" style="font:arial" onload="document.getElementById('select').focus()">
<%
response.Write("<form method='post' action='noventa_t.asp?vend=" & replace(vend,"Ñ","NN") & "&cliente=" & cliente& "'>")
%>
<table width="100%" border="0">
<tr>
  <td height="29" align="center" valign="middle"><strong><FONT SIZE="2" face="verdana" COLOR="#000066">MOTIVO VENTA NO EFECTUADA</FONT></strong></td>
<tr>
    <td height="37" align="center">
		<select name="select" id="select" size="0" onChange="textbox_onChange()">
			<option value=""                 >SELECCIONE       </option>
			<option value="Tiene Stock"      >Tiene Stock      </option>
			<option value="Encargado Ausente">Encargado Ausente</option>
			<option value="CtaCte Bloqueada" >CtaCte Bloqueada </option>
			<option value="Fono no contesta" >Fono no contesta </option>
			<option value="Otro"             >Otro             </option>
        </select>
	</td>
  </tr>
  <tr align="center" valign="middle">
    <td height="34"><strong>Observaciones</strong> : </td>
  </tr>
  <tr align="center" valign="middle">
    <td height="67"><textarea name="textarea" id="textarea"  cols="25" rows="3" readonly="true"></textarea></td>
  </tr>
  <tr>
    <td height="45" align="center" valign="middle">
	<INPUT TYPE="button" value="<< Volver" onclick="history.back()">&nbsp;&nbsp;
	<input name="enviar" id="enviar" type="submit" value="Enviar" style="display:'none'">
      </td>
    </tr>
</table>
</form></br>
</body>
</html>

<SCRIPT LANGUAGE="JavaScript">
<!--
function textbox_onChange(){
	var objeto='select';
	mivalor=document.getElementById(objeto).value;
	if (mivalor==''){
		document.getElementById('enviar').style.display='none';
	}else{
		document.getElementById('enviar').style.display='';
		document.getElementById('enviar').focus()
		//document.formulario2.submit();
	};

	if (mivalor=='Otro'){
		document.getElementById('textarea').removeAttribute('readOnly'); 
		document.getElementById('textarea').focus()
	}else{
		document.getElementById('textarea').readOnly=true
	}

	//alert(mivalor);
	
}
//-->
</SCRIPT>