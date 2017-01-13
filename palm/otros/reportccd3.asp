<% Function ProperCase(sString) 
  ' Declare the variables 
  Dim sWhiteSpace, bCap, iCharPos, sChar 

  ' Valid white space characters. Add more as necessary
  sWhiteSpace = Chr(32) & Chr(9) & Chr(13) 

  ' Convert the string to lowercase, to start with 
  sString = LCase(sString)

  ' Initialize the caps flag to True 
  bCap = True 
  
  ' Loop through each of the characters in the string 
  For iCharPos = 1 to Len(sString) 
    ' Get one character at a time
    sChar = Mid(sString, iCharPos, 1) 

    ' If caps flag is true, upper case the character 
    If bCap = True Then 
      sChar = UCase(sChar) 
    End If 

    ' Append to the final string 
    ProperCase = ProperCase + sChar 

    ' If the character is a white space, raise the caps flag 
    If InStr(sWhiteSpace, sChar) Then 
      bCap = True 
    Else 
      bCap = False 
    End If 
  Next 
End Function
%><html>
<head>
<title>Reporte CCD Por Vendedor</title>
<style type="text/css">
<!--
#bordes {
	border:thin outset;
}
#vensel {
	text-transform: capitalize;
	background-color: #CCCCCC;
	font-weight: bold;
}
-->
</style>
</head>
<body>
<form method="post" action="reportccd3b.asp">
<table align="center" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC" id="bordes">
  <tr>
	<td colspan="3">&nbsp;</td>
	</tr>
  <tr>
	<td>&nbsp;</td>
	<td align="center"><strong>Vendedor:</strong></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<td align="center">
	<select name="Vendedor" id="vensel"><%
set oConn=server.createobject("ADODB.Connection")
oConn.open "Provider=SQLOLEDB;Data Source=serverdesa;Initial Catalog=BDFlexline;User Id=sa;Password=desakey;"

Sql="SELECT RIGHT('000' + CAST(ID_Vendedor AS nvarchar), 3) AS NU_2, Vendedor AS Nombre "&_
    "FROM Flexline.PDA_CCD_DOC WHERE (prep = 0) "&_
    "GROUP BY Vendedor, RIGHT('000' + CAST(ID_Vendedor AS nvarchar), 3) "&_
    "ORDER BY RIGHT('000' + CAST(ID_Vendedor AS nvarchar), 3)"
set rs1=oConn.execute(Sql)
do until rs1.eof
%>
			<option value="<%= rs1.fields("Nombre")%>">
				<%= rs1.fields("NU_2") & " - " & lcase(replace(rs1.fields("Nombre"),"Ñ","n"))%>			</option>
			<%
rs1.movenext
loop
%>
		</select></td>
  	<td>&nbsp;</td>
  </tr>
  <tr>
	<td colspan="3">&nbsp;</td>
	</tr>
  <tr>
	<td colspan="3">&nbsp;</td>
	</tr>
  <tr>
	<td colspan="3">&nbsp;</td>
	</tr>
  <tr>
	<td align="center">&nbsp;</td>
	<td align="center"><input type="submit" value="Mostrar Pagos"></td>
    <td align="center">&nbsp;</td>
  </tr>
	<tr>
	<td colspan="3"><input type="hidden" name="BT" value="SGTE">&nbsp;</td>
	</tr>
</table>
</form>
</body>
</html>
