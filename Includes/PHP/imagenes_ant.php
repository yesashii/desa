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
    $size = '170';

?>

		<div style="position:relative;width:600px;height:230px;" border="1">
			<div style="position:absolute;left:0px;top:0px;width:600px;height:200px;z-index:111;opacity:0.5;filter:alpha(opacity=50);">
				<!-- <IMG SRC="http://distilleryimage6.s3.amazonaws.com/eab9c79a13e211e3815e22000a1f8e11_6.jpg" WIDTH="407" HEIGHT="415" BORDER="0" ALT=""> -->
				<TABLE width="100%" height="100%" bgcolor="#003366" border="0">
				<TR>
					<TD></TD>
				</TR>
				</TABLE>
			</div>
			<div style="overflow-x:scroll;overflow-y:hidden; position:absolute;left:0px;top:0px;width:600px;height:230px;z-index:333;">
				<TABLE width="100%" height="200px" border="0">
				<TR>
<?php
	// Show results ;overflow:hidden
    // Using for loop will cause error if there are less photos than the limit
    foreach(array_slice($media->data, 0, $limit) as $data)
    {
        // Show photo
$imagen=$data->images->low_resolution->url;
if ($imagen=='http://distilleryimage1.s3.amazonaws.com/f1b0c9b4171511e3ae5e22000a1f8fb9_6.jpg'){$imagen='';};
if ($imagen=='http://distilleryimage6.s3.amazonaws.com/eab9c79a13e211e3815e22000a1f8e11_6.jpg'){$imagen='';};
if ($imagen=='http://distilleryimage0.s3.amazonaws.com/274f58a61d6011e3af0122000a1fbc9e_6.jpg'){$imagen='';};
if ($imagen!=''){
		echo '<TD valign="center" align="center"><img src="'.$imagen.'" height="'.$size.'" width="'.$size.'" alt=""></TD>';	
		}
    }
//Fin Instagram

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
if ($imagen=='https:\/\/pbs.twimg.com\/media\/BTknAP7IQAAG1Rw.jpg'){$imagen='';};
if ($imagen=='https:\/\/pbs.twimg.com\/media\/BTbZ1uhIcAAlhPR.jpg'){$imagen='';};
if ($imagen!=''){
		//echo '<TD valign="center" align="center"><img src="'.$imagen.'" height="'.$size.'" width="'.$size.'" alt=""></TD>';	
		//echo $imagen;
		}
	}
}
//twiter a mano
echo '<TD valign="center" align="center"><img src="https://pbs.twimg.com/media/BUJBc1DCUAAeZbI.jpg" height="'.$size.'" width="'.$size.'" alt=""></TD>';	
echo '<TD valign="center" align="center"><img src="https://pbs.twimg.com/media/BUJmgoBIQAA7kas.jpg" height="'.$size.'" width="'.$size.'" alt=""></TD>';
echo '<TD valign="center" align="center"><img src="https://pbs.twimg.com/media/BULOdjaIMAA2Skd.jpg" height="'.$size.'" width="'.$size.'" alt=""></TD>';
 //       $result .= '\t'.'<a title="'.htmlentities($value->caption->text).' ('.htmlentities(date("d/m/Y", $value->caption->created_time)).')" href="'.$value->images->standard_resolution->url.'">';
 //          $result .= '<img src="'.$value->images->low_resolution->url.'" alt="'.$value->caption->text.'" width="'.$width.'" height="'.$height.'" />';
 //         $result .= '</a>';
//----------------------------------------------------------------------------------------------
function imgdx ($iruta ='', $iuser = ''){ 
   echo "a = ".$a."<br/>"; 
   echo "b = ".$b."<br/>"; 
   echo "<br/>"; 
} 
//----------------------------------------------------------------------------------------------
function default_values_test ($a = 123, $b = 456){ 
   echo "a = ".$a."<br/>"; 
   echo "b = ".$b."<br/>"; 
   echo "<br/>"; 
} 
//----------------------------------------------------------------------------------------------
?>
				</TR>
				</TABLE>
			</div>
		</div>


 <!-- height="'.$size.'" width="'.$size.'" -->
 <!-- low_resolution -->
 <!-- standard_resolution -->


 <?php


// Perform the request
//$twitter = new TwitterAPIExchange($settings);
//$media = $instagram->getTagMedia($tag);

// Make the request and get the response into the $json variable


//echo '<HR>';

//}

//-------------------------------------------------------------
?>