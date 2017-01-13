<% Dim oConn, Vnd
set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open "Provider=SQLOLEDB;Data Source=sqlserver;Initial Catalog=desaerp;User Id=sa;Password=desakey"
nuser = Request.Querystring("nuser")
np    = Request.Querystring("np")
%>
<html>
<head>
<title>Editar Pedido</title>
</head>
<body>
<%
sql="update PED_PEDIDOSENC " & _
"set sw_estado='N' " & _
"WHERE     (numero_pedido = N'" & np & "') and SW_estado in('I','P')"
Set rs = oConn.Execute(sql)
'response.write(sql)
response.write("<BR><BR>Pedido Eliminado : " & np)
'-------
function iif(bolean,vtrue,vfalse)
	if bolean then
		iif = vtrue
	else
		iif =vfalse
	end if
end function
%>
<FORM METHOD=POST ACTION="editpedido.asp?nuser=<%=nuser%>">
<INPUT TYPE="submit" value="Volver">
</FORM>

<FORM METHOD=POST ACTION="../../default.asp?nuser=<%=nuser%>">
<INPUT TYPE="submit" value="Inicio">
</FORM>
</body>
