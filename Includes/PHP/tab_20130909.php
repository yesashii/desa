<!DOCTYPE html>
<html xmlns:fb="http://www.facebook.com/2008/fbml">
<HEAD>
<TITLE>savory</TITLE>
<META NAME="Generator"   CONTENT="EditPlus">
<META NAME="Author"      CONTENT="Simon Hernandez">
<META NAME="Keywords"    CONTENT="">
<META NAME="Description" CONTENT="app facebook">
<style type="text/css">
	#noprint {
		display:none;
	}
	body{
		background-color: #FFFFFF;
		font-family: verdana;
		font-size: 12px;
	}
	th{
		font-family: verdana;
		font-size: 12px;
		color: #000033;
		padding: 2px;
		/*border-style: solid;
		border-width: 2px;
		background-color: #000066*/
	}
	td{
		font-family: verdana;
		font-size: 12px;
		padding: 2px;
		/*border-style: solid;
		border-width: 2px;
		padding: 2px;*/
	}
	table{
		border-collapse: collapse;
		/*border: 1px;
		*border-color: #000066;
		background-color: #FFFFFF;*/
	}
</style>
</HEAD>

<script>
function myLogin(){
  var cb = function(response) {
    Log.info('FB.login callback', response);
    if (response.status === 'connected') {
      Log.info('User logged in');
	  //window.location.reload();
    } else {
      Log.info('User is logged out');
    }
  };
  FB.login(cb, {
  // establece permisos a solicitar
    scope: 'publish_stream',
    enable_profile_selector: 1
		//window.location.reload();
  });
//  FB.ui(
	//
	//location.reload();
	//pagina='tab.php?v=1';
	//location.href=pagina;
	FB.Event.subscribe('auth.login', function(response) {
          //window.location.reload();
		  //https://www.facebook.com/savoryprontocontigo/app_618439011509777
		  //window.location.href = 'https://www.facebook.com/savoryprontocontigo/app_618439011509777';
		  //setTimeout ("window.location.href = 'https://www.facebook.com/savoryprontocontigo/app_618439011509777?V=1'", 10000);
		  setTimeout ("window.location.href =tabz.php'", 10000);
        });
	FB.Event.subscribe('auth.logout', function(response) {
          //window.location.reload();
		  //window.location.href = 'https://www.facebook.com/savoryprontocontigo/app_618439011509777';
		  //setTimeout ("window.location.href = 'https://www.facebook.com/savoryprontocontigo/app_618439011509777'", 10000);
		  setTimeout ("window.location.href =tabz.php'", 10000);
        });
};
</script>

  <body>
<?php
// .:: INCLUDES ::.
require_once('appinclude.php');


flush();

//$like_status='1';
$signed_request = $facebook->getSignedRequest();
$like_status = $signed_request["page"]["liked"];

//echo $like_status;


// Si el usuario le ha clickado en el "Me gusta" de nuestra p�gina
if($like_status){
	// Ver si hay un usuario de una cookie
	$user = $facebook->getUser();

	if ($user) {
	  try {
		// Continuar sabiendo que tenemos un usuario conectado que est� autenticado.
		$user_profile = $facebook->api('/me');
	  } catch (FacebookApiException $e) {
		echo '<pre>'.htmlspecialchars(print_r($e, true)).'</pre>';
		$user = null;
	  }
	}

	if ($user) { ?>


<script language="JavaScript" type="text/javascript">

//var pagina="galeria.php"
function redireccionar() 
{
//location.href=pagina
	document.getElementById('dv_video').style.display='none';
	document.getElementById('dv_ganadores').style.display='none';
	document.getElementById('dv_galeria').style.display='';
	//alert(document.getElementById('videoplay').src);
	document.getElementById('videoplay').src='http://www.youtube.com/embed/J5VzczoH7Sw?autoplay=1'; 
} 
function showganadores() 
{
//location.href=pagina
	document.getElementById('dv_video').style.display='none';
	document.getElementById('dv_galeria').style.display='none';
	document.getElementById('dv_ganadores').style.display='';
} 
setTimeout ("redireccionar()", 50000);

</script>
              
<CENTER>
<!-- Div video -->
<DIV id="dv_video">
<TABLE width="800px" background="../sabory_201309/img/layer1.jpg" border="0">
<TR height="187">
	<TD width="5px"></TD>
	<TD></TD>
</TR>
<TR>
	<TD></TD>
	<TD align="center">
 	<iframe id="videoplay" class="youtube-player" type="text/html" width="426px" height="240px" src="http://www.youtube.com/embed/R5U0kTKdf5k?autoplay=1" frameborder="0"></iframe> 


<BR>
	<!-- boton saltar -->
	<TABLE width="400px" border="0">
	<TR>
		<TD align="right"><a id="myLink" href="#" onclick="redireccionar();return false;"><IMG SRC="../sabory_201309/img/bot_saltar.png" WIDTH="120" HEIGHT="52" BORDER="0" ALT=""></A></TD>
	</TR>
	</TABLE>
	</TD>
</TR>
<TR height="104">
	<TD></TD>
	<TD></TD>
</TR>
</TABLE>
</DIV>
<!-- Div Galeria -->
<DIV id="dv_galeria" style="display:none">
<TABLE border="0" width="100%" background="../sabory_201309/img/layer2.jpg">
<TR height="280px">
	<TD width="180px"></TD>
	<TD></TD>
</TR>
<TR>
	<TD align="left" valign="top">
		<BR><BR>
		<a id="myLink" href="#" onclick="redireccionar();return false;"><IMG SRC="../sabory_201309/img/bot_concurso.png"  WIDTH="121" HEIGHT="52" BORDER="0" ALT="Galeria"></A>
		<a id="myLink" href="#" onclick="showganadores();return false;"><IMG SRC="../sabory_201309/img/bot_ganadores.png" WIDTH="121" HEIGHT="52" BORDER="0" ALT="Ganadores"></A>
		<A HREF="bases.pdf"     target="_blank" ><IMG SRC="../sabory_201309/img/bot_bases.png"     WIDTH="121" HEIGHT="52" BORDER="0" ALT="Bases"></A>
	</TD>
	<TD align="center">
		<iframe src="imagenes.php" id="galeria" width="100%" height="300" frameborder="0" scrolling="no">#HeladosaLaParrilla</iframe>
	</TD>
</TR>
<TR height="10px">
	<TD></TD>
	<TD></TD>
</TR>
</TABLE>
<!-- 
      Your user profile is
      <pre> -->
        <!-- php print htmlspecialchars(print_r($user_profile, true)) --> 

<!-- inicio comentario -->
<!-- <div class="fb-comments" data-href="https://www.socialmediaconsulting.cl/app_facebook/sabory_201309/" data-width="470"></div> --><!-- comentarios -->
<div id="fb-root"></div>
<script>(function(d, s, id) {
  var js, fjs = d.getElementsByTagName(s)[0];
  if (d.getElementById(id)) return;
  js = d.createElement(s); js.id = id;
  js.src = "//connect.facebook.net/es_LA/all.js#xfbml=1&appId=618439011509777";
  fjs.parentNode.insertBefore(js, fjs);
}(document, 'script', 'facebook-jssdk'));</script>
<!-- fin comentario -->
</DIV>

<!-- Div Ganadores -->
<DIV id="dv_ganadores" style="display:none">
<TABLE border="0" width="100%" background="../sabory_201309/img/layer3.jpg">
<TR height="280px">
	<TD width="180px"></TD>
	<TD></TD>
</TR>
<TR  height="300px">
	<TD align="left" valign="top">
		<BR><BR>
		<a id="myLink" href="#" onclick="redireccionar();return false;"><IMG SRC="../sabory_201309/img/bot_concurso.png"  WIDTH="121" HEIGHT="52" BORDER="0" ALT="Galeria"></A>
		<a id="myLink" href="#" onclick="showganadores();return false;"><IMG SRC="../sabory_201309/img/bot_ganadores.png" WIDTH="121" HEIGHT="52" BORDER="0" ALT="Ganadores"></A>
		<A HREF="bases.pdf"     target="_blank" ><IMG SRC="../sabory_201309/img/bot_bases.png"     WIDTH="121" HEIGHT="52" BORDER="0" ALT="Bases"></A>
	</TD>
	<TD align="center">
		<!-- IMAGENES GANADORES -->
	</TD>
</TR>
<TR height="10px">
	<TD></TD>
	<TD></TD>
</TR>
</TABLE>
</DIV>
</CENTER>



<script>
FB.init({
appId:'<?php echo $facebook->getAppID() ?>',
cookie:true,
status:true,
xfbml:true

});
 
function FacebookInviteFriends()
{
FB.ui({
method: 'apprequests',
message: ' te invita al fanPage de Savory'

,redirect_uri:'https://www.facebook.com/SocialAdvisors'
,title: 'Te invita al fanPage de Facebook',
//link: '**LINK DE NUESTRA P�GINA FACEBOOK**',
//picture: '**LINK A UNA IMAGEN QUE SE MOSTRAR� EN EL DI�LOGO',
//caption: '**MENSAJE DEL DI�LOGO',
//description: '**TEXTO CON LA DESCRIPCI�N DEL DI�LOGO** '

});
}
</script>
<!--  
<div id="fb-root"></div>
 
<a href='#' onclick="FacebookInviteFriends();"> 
<BR>Invita a tus Amigos de Facebook!</a> -->
<!-- </pre> -->

    <?php } else { 
		print_paso01();
		?>
		<SCRIPT LANGUAGE="JavaScript">
		<!--
		//myLogin();
		//-->
		</SCRIPT>
		<?php

	} //end if
	?><div id="fb-root"></div><?php

	//echo $user;
	
}else{
	print_hazmegusta();
}

// call pie de pagina
escribe_pie("Copyright 2013. Social Advisors. Todos los derechos reservados");


//-----------------------------------------------------------------------
function escribe_pie($cadena){
	?>	
	<CENTER><FONT SIZE="1" face='verdana' COLOR="#FFFFFF">
		<?php echo $cadena ;?>
	</FONT></CENTER>
	<?php
}
//-----------------------------------------------------------------------
function print_hazmegusta(){


	//echo $like_status;
	?>
 	<TABLE width="100%">
	<TR>
		<TD align="right"><IMG SRC="http://pda.desa.cl/includes/php/hazclick.jpg" BORDER="0" ALT="Haz click en me gusta" onclick="window.location.reload()"></TD>
	</TR>
	</TABLE>
	<BR><BR><BR><BR> 
	<!-- <iframe src="layer0.swf?=1" id="galeria" width="800" height="600" frameborder="0" scrolling="no">Sabory</iframe> -->
	<?php
}
//-----------------------------------------------------------------------
function print_paso01(){
	?>
	<CENTER>
	<div id="psaso0">
		<TABLE border="0" width="800px" height="600px" >
		<TR>
			<TD><IMG SRC="../sabory_201309/img/layer0.jpg" WIDTH="800" HEIGHT="600" BORDER="0" ALT="" onclick="myLogin()"></TD>
		</TR>
		</TABLE>
			<div style="position:absolute;left:0px;top:0px;width:10px;height:10px;z-index:111;opacity:0;filter:alpha(opacity=0);">
			<img id="fb-login" src="../sabory_201309/img/btn_aceptar1.PNG"
			onclick="myLogin()"
			style="cursor:pointer"/>
			</div>
		</div>
	</CENTER>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	setTimeout ("myLogin()", 1000);
	//alert('o');
	//-->
	</SCRIPT>
	<?php
}
//-----------------------------------------------------------------------
//-----------------------------------------------------------------------
//-----------------------------------------------------------------------

?>
    <script>
      window.fbAsyncInit = function() {
        FB.init({
          appId: '<?php echo $facebook->getAppID() ?>',
          cookie: true,
          xfbml: true,
          oauth: true
        });
        FB.Event.subscribe('auth.login', function(response) {
          window.location.reload();
        });
        FB.Event.subscribe('auth.logout', function(response) {
          window.location.reload();
        });
      };
      (function() {
        var e = document.createElement('script'); e.async = true;
        e.src = document.location.protocol +
          '//connect.facebook.net/en_US/all.js';
        document.getElementById('fb-root').appendChild(e);
      }());
    </script>
	<!-- Script Log In -->


  </body>
</html>