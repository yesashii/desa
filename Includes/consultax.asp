<% 
Set oConnx = server.createobject("ADODB.Connection")
oConnx.open "Provider=SQLOLEDB;Data Source=SQLSERVER;Initial Catalog=todo;User Id=sa;Password=desakey;"

sql = trim(cstr(request.form("sql")))
if len(sql)=0 then sql=request.querystring("sql")
if len(sql)=0 then sql="select top 10 * from serverdesa.BDFlexline.flexline.producto"

sql=replace(sql,"%20"," ")
sql=replace(sql,":","=")
'response.write sql
sql=replace(sql,"likex","%")
'on error resume next
Set rs=oConnx.execute(Sql)
if rs.eof then
	response.write "No encontrado"
else
	response.write cisnull(rs(0),"")
end if

rs.close
oConnx.close
'------------------------------------------
function cisnull(valor, alternativo)
	if isnull(valor) then
		cisnull=alternativo
	else
		cisnull=valor
	end if
end function
%>