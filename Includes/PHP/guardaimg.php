<?php
$img=htmlspecialchars($_POST['idimagen']); 
$acc=htmlspecialchars($_POST['accion']);
$usr=htmlspecialchars($_POST['user']);
$org=htmlspecialchars($_POST['origen']);

echo ''.$img;
echo '<BR>'.$acc;
echo '<BR>'.$usr;
echo '<BR>'.$org;
echo '<HR>';

//conecta();
	$conexion = mysqli_connect("localhost","socialme_admin","qwe123","socialme_savory");
	//"localhost",socialmediaconsulting.cl
	if (mysqli_connect_errno())
	{
		echo "Failed to connect to MySQL: " . mysqli_connect_error();
	}

//mysql_select_db("socialme_savory", $conexion); 
//$consulta =  mysqli_query("SELECT * FROM imagenes LIMIT 0,30",$conexion);
$consulta ="SELECT * FROM imagenes LIMIT 0,30";

if ($resultado = mysqli_query($conexion, $consulta)) {
	//mysqli_query($enlace, $consulta)

    /* obtener el array asociativo */
    while ($fila = mysqli_fetch_row($resultado)) {
        //printf ("%s (%s)\n", $fila[0], $fila[1]);
		echo '<BR>I:'.$fila[1];
    }

    /* liberar el conjunto de resultados */
    mysqli_free_result($resultado);
}


//echo $consulta;
//$link = mysql_connect("localhost", "nobody"); 
//mysql_select_db("mydb", $link); 
//$result = mysql_query("SELECT nombre, email FROM agenda", $link); 
//echo "<table border = '1'> \n"; 
//echo "<tr><td>Nombre</td><td>E-Mail</td></tr> \n"; 
//while ($row = mysqli_fetch_row($consulta)){ 
//       echo "<tr><td>$row[0]</td><td>$row[1]</td></tr> \n"; 
//} 
//echo "</table> \n"; 

//mysqli_close($conexion);


//conecta();

if ($acc=='activar'){
	//echo '1';
	$sql="INSERT INTO imagenes (user,img,vigencia,origen) VALUES('".$usr."','".$img."',1,'".$org."');";
	echo $sql;
	//mysqli_query($sql);
	if (mysqli_query($conexion, $sql)) {
		print "<p>Registro a�adido correctamente.</p>";
	} else {
		print "<p>Error al a�adir el registro.</p>";
		//echo mysqli_error();
	}
}

if ($acc=='upactivar'){
	//echo '1';
	$sql="update imagenes set vigencia='1' where img='".$img."';";
	echo $sql;
	//mysqli_query($sql);
	if (mysqli_query($conexion, $sql)) {
		print "<p>Registro Actualizado correctamente.</p>";
	} else {
		print "<p>Error al Actualizado el registro.</p>";
		//echo mysqli_error();
	}
}

if ($acc=='rechazar'){
	$sql="INSERT INTO imagenes (user,img,vigencia,origen) VALUES('".$usr."','".$img."',0,'".$org."');";
	if (mysqli_query($conexion, $sql)) {
		print "<p>Registro a�adido correctamente.</p>";
	} else {
		print "<p>Error al a�adir el registro.</p>";
	}
//echo $sql;
}

if ($acc=='uprechazar'){
	$sql="update imagenes set vigencia='0' where img='".$img."';";
	if (mysqli_query($conexion, $sql)) {
		print "<p>Registro a�adido correctamente.</p>";
	} else {
		print "<p>Error al a�adir el registro.</p>";
	}
//echo $sql;
}

/*
$con=mysqli_connect("example.com","peter","abc123","my_db");
// Check connection
if (mysqli_connect_errno())
  {
  echo "Failed to connect to MySQL: " . mysqli_connect_error();
  }

mysqli_query($con,"INSERT INTO Persons (FirstName, LastName, Age)
VALUES ('Peter', 'Griffin',35)");

mysqli_query($con,"INSERT INTO Persons (FirstName, LastName, Age) 
VALUES ('Glenn', 'Quagmire',33)");

mysqli_close($con);

*/
mysqli_close($conexion);

//-----------------------------------------------------------------------------
function conecta(){ 
	$conexion = mysqli_connect("localhost","socialme_admin","qwe123","socialme_savory");
	//"localhost",socialmediaconsulting.cl
	if (mysqli_connect_errno())
	{
		echo "Failed to connect to MySQL: " . mysqli_connect_error();
	}
} 
//-----------------------------------------------------------------------------

function ejecutar_sql($sql){
	$resultado = mysql_query($sql);
	if (! $resultado ) {die("ERROR AL EJECUTAR LA CONSULTA: ".mysql_error());}
	return $resultado;
}

?> 