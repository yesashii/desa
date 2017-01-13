<html>
<head>
<title>Traspaso</title>
</head>
<body bgcolor="#FFFFFF">
pedido
<%
'-----------------------------------------------------------------------------
sub guardaped()
':::::::::::::::::: conexion :::::::::::::::::
private tipodoc, mitotal, oConn, rs, sql, rs1, SQL2, oConn2
Set oConn  = server.createobject("ADODB.Connection")
oConn.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=todo;User Id=sa;Password=desakey;"

MiID=request.querystring("id")
End sub 'guardaped()
'-----------------------------------------------------------------------------
%>
</body>
</html>