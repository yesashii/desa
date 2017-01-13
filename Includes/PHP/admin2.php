<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
</HEAD>

<BODY>
<?php 
    // Get class for Instagram
    // More examples here: https://github.com/cosenary/Instagram-PHP-API
    require_once 'instagram.class.php';

    // Initialize class with client_id
    // Register at http://instagram.com/developer/ and replace client_id with your own
    $instagram = new Instagram('15f3f78de4c74544ab78b4c3908e70ac');

    // Set keyword for #hashtag
    $tag = 'heladosalaparrilla';

    // Get latest photos according to #hashtag keyword
    $media = $instagram->getTagMedia($tag);

    // Set number of photos to show
    $limit = 5;

    // Set height and width for photos
    $size = '300';

$conexion = mysql_connect("localhost","socialme_admin","qwe123");
?>
<iframe src="null.php" id="galeria" width="100%" height="60" frameborder="0" scrolling="no">#HeladosaLaParrilla</iframe>
<TABLE>
<?php

		// Show results ;overflow:hidden
    // Using for loop will cause error if there are less photos than the limit
    foreach(array_slice($media->data, 0, $limit) as $data)
    {
        // Show photo
$imagen=$data->images->low_resolution->url;
//if ($imagen=='http://distilleryimage1.s3.amazonaws.com/f1b0c9b4171511e3ae5e22000a1f8fb9_6.jpg'){$imagen='';};
//if ($imagen=='http://distilleryimage6.s3.amazonaws.com/eab9c79a13e211e3815e22000a1f8e11_6.jpg'){$imagen='';};
if ($imagen!=''){
		echo '<TR>';
		echo '<TD><img src="'.$imagen.'" height="'.$size.'" width="'.$size.'" alt=""></TD>';
		echo '<TD>Instagram</TD>';
		echo '</TR>';
		}
    }







//INICIO Twitter
ini_set('display_errors', 1);
require_once('TwitterAPIExchange.php');

/** Set access tokens here - see: https://dev.twitter.com/apps/ **/
$settings = array(
    'oauth_access_token' => "1739888882-Z1QrRm5QJszHnweUNVJwVi9fCPr9AjBFXle02JY",
    'oauth_access_token_secret' =>  "cEbItQHqEYUCnfpH9ceRMLIgspjwri2vy49OSORvk",
    'consumer_key' =>  "4tMjmmzfV36MvIKz4EwU6Q",
    'consumer_secret' =>  "TcIFif3EPCPVbiam4cJvD9eEl8dj5MD6cnz1mjHQo4"

);
$url = 'https://api.twitter.com/1.1/search/tweets.json';
$requestMethod = 'GET';
$getfield = '?q=#heladosalaparrilla&result_type=recent';
// Perform the request
$twitter = new TwitterAPIExchange($settings);
$txt = $twitter->setGetfield($getfield)->buildOauth($url, $requestMethod)->performRequest() ;
$a=explode('"metadata"',$txt);
foreach ($a as &$valor) {
	$string = strstr($valor, 'media_url'); 
	$e=explode('","',$string);
	$e[1] = str_replace('media_url_https":"', '', $e[1] );
	if (strlen($e[1])>0){
		//echo '<HR>'.$e[1];
		$imagen=$e[1];
//if ($imagen=='https:\/\/pbs.twimg.com\/media\/BTknAP7IQAAG1Rw.jpg'){$imagen='';};
//if ($imagen=='https:\/\/pbs.twimg.com\/media\/BTbZ1uhIcAAlhPR.jpg'){$imagen='';};
if ($imagen!=''){
		echo '<TD valign="center" align="center"><img src="'.$imagen.'" height="'.$size.'" width="'.$size.'" alt=""></TD>';	
		echo $imagen;
		}
	}
}

?>
</TABLE>
<?php

//$conexion = mysql_connect("localhost","socialme_admin","qwe123");


//selec
$consulta = mysql_query("SELECT * FROM imagenes ) or die ('Error en la consulta'");

while($fila=mysql_fetch_array($consulta)){
$user = $fila['user'];
$img = $fila['img'];
$correo = $fila['correo'];
echo "<p>E producto $user ($correo) vale $ $img</p>";
// se obtienen multiples párrafos variando con los datos de los distintos productos 
}

/*

//guardar
$conexion = mysql_connect("localhost","socialme_admin","qwe123");

mysql_select_db("socialme_savory",$conexion);

$sql="INSERT INTO imagenes(user,img,vigencia,check) VALUES('','',0,0)";

mysql_query($sql);
*/
//---------------------------------------------------
function existe($miimg)
{
    //return array (0, 1, 2);
	$consulta = mysql_query("SELECT * FROM imagenes ) or die ('Error en la consulta'");
}
?>
</BODY>
</HTML>
