<html>
<head>
<title>Traspaso</title>
</head>
<body bgcolor="#FFFFFF">
<FONT SIZE="1">pedido 
<%
response.write(request.querystring("id"))
private tipodoc, mitotal, oConn, rs, sql, rs1, SQL2, oConn2
Set oConn  = server.createobject("ADODB.Connection")
oConn.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=todo;User Id=sa;Password=desakey;"
MiID=request.querystring("id")

call guardaped()
'-----------------------
sub guardaped()
	if len(MiID)=0 then
		exit sub
	end if

	sql="select * FROM todo.flexline.FX_PEDIDO_PDA FX_PEDIDO_PDA where id=" & MiID
	Set rs=oConn.execute(Sql)
	
	if rs.eof then
		exit sub
	end if
on error resume next

	rs.Close
	rs.Open SQL, oConn, 1, 3
	
	For Each elemento In request.querystring
		if elemento<>"id" then
			if elemento="minota" then
				rs.fields("nota")=Request.querystring(elemento)
			else
				if elemento <>"fechaentrega" then 
					rs.fields(elemento)=Request.querystring(elemento)
				else
					if len(Request.querystring(elemento))>0 then rs.fields(elemento)=Request.querystring(elemento) 'fecha entrega
				end if
			end if

			Response.Write("<BR>" & elemento & " : " & Request.querystring(elemento))
			
		end if
	Next
	rs.Update

	x = 1

End sub 'guardaped()
'-----------------------
%></FONT>
</body>
</html>