<?php
//::conexion SQL
	$conexionsql = mysqli_connect("localhost","socialme_admin","qwe123","socialme_savory");
	//echo $conexion;
	//"localhost",socialmediaconsulting.cl
	if (mysqli_connect_errno())
	{
		echo "Failed to connect to MySQL: " . mysqli_connect_error();
	}


require_once('twitteroauth.php');
 
$consumerKey    = '4tMjmmzfV36MvIKz4EwU6Q';
$consumerSecret = 'TcIFif3EPCPVbiam4cJvD9eEl8dj5MD6cnz1mjHQo4';
$oAuthToken     = '1739888882-Z1QrRm5QJszHnweUNVJwVi9fCPr9AjBFXle02JY';
$oAuthSecret    = 'cEbItQHqEYUCnfpH9ceRMLIgspjwri2vy49OSORvk';
 
$conexion = new TwitterOAuth($consumerKey, $consumerSecret, $oAuthToken, $oAuthSecret);


//    'oauth_access_token' => "1739888882-Z1QrRm5QJszHnweUNVJwVi9fCPr9AjBFXle02JY",
//    'oauth_access_token_secret' =>  "cEbItQHqEYUCnfpH9ceRMLIgspjwri2vy49OSORvk",
//    'consumer_key' =>  "4tMjmmzfV36MvIKz4EwU6Q",
//    'consumer_secret' =>  "TcIFif3EPCPVbiam4cJvD9eEl8dj5MD6cnz1mjHQo4"

//require_once('datos_oAuth.php');
//$tweets = $conexion->get("https://api.twitter.com/1.1/LO_QUE_HAREMOS_EN_URL");
//$tweets = $conexion->get("https://api.twitter.com/1.1/statuses/user_timeline.json?screen_name=carlitoxenlaweb&count=5");
$tweets = $conexion->get("https://api.twitter.com/1.1/search/tweets.json?q=%23heladosalaparrilla&result_type=recent&count=2000");
//https://api.twitter.com/1.1/search/tweets.json?q=%23freebandnames&since_id=24012619984051000&max_id=250126199840518145&result_type=mixed&count=4
//?q=#heladosalaparrilla&result_type=recent
 
// Creamos el JSON
$json = json_encode($tweets);
//echo $json;
 
echo "<br/><br/><hr/>Ahora mas entendible...<hr/><br/>";
 
// Procesamos el JSON
$arrJson = json_decode($json, true);


//echo "<pre>";
  //print_r($arrJson[statuses]);
//echo "</pre><HR>";


echo '<div>';
//foreach ($arrJson as $key=>$valor) {
foreach ($arrJson[statuses] as &$valor){
	
if ($valor[entities][media][0][media_url]){
	echo "<pre>";
		print_r($valor);
	echo "</pre><HR>";

	echo '<HR><IMG SRC="'.$valor[user][profile_image_url].'" ><IMG SRC="'.$valor[entities][media][0][media_url].'" width="100" height="100" >';
	echo '<BR>@'.$valor[user][screen_name];
	echo '<BR>'.$valor[user][name];
	echo '<BR>'.$valor[text] ;
	$img=$valor[entities][media][0][media_url_https] ;
	echo '<BR>'.$img ;
	
	echo '<BR>existe: '.existe($img, $conexionsql);
	if (existe($img, $conexionsql)=="1"){
		$sql="update imagenes set correo='".$valor[created_at]."', user='@".$valor[user][screen_name]."' where img='".$img."';";
		echo $sql;
		if (mysqli_query($conexionsql, $sql)) {
			print "<p>Registro a�adido correctamente.</p>";
		} else {
			print "<p>Error al a�adir el registro.</p>";
		}
	}
	//echo '<BR><IMG SRC="'.$valor[entities][media][0][media_url].'" width="100" height="100" >';
	}
}
echo '</pre>';


//---------------------------------------------------
function existe($miimg, $oconn){
    //return array (0, 1, 2);
	//$consulta = mysql_query("SELECT * FROM imagenes ) or die ('Error en la consulta'");
	$consulta ="SELECT img FROM imagenes where img ='".$miimg."' LIMIT 0,30";
	//$consulta ="SELECT img FROM imagenes LIMIT 0,30";
	echo '<HR>'.$consulta;
	//echo $conexion;
	$respuesta ="0";
	if ($resultado = mysqli_query($oconn, $consulta)) {
		while ($fila = mysqli_fetch_row($resultado)) {
		//$respuesta =$fila[0];
		$respuesta="1";
		echo '<HR>I:'.$fila[0];
		}
		mysqli_free_result($resultado);
	}
	return $respuesta;
}
//---------------------------------------------------
mysqli_close($conexionsql);

?>
