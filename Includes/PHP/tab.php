<?php
require_once('facebook.php');

// Iniciamos nuestra APP
$facebook = new Facebook(array(
  'appId'  => '618439011509777',
  'secret' => '59d3f7477560674f8d3c49cc9767fad3',
    'cookie' => true,
));

$signed_request = $facebook->getSignedRequest();
$like_status = $signed_request["page"]["liked"];

// Si el usuario le ha clickado en el "Me gusta" de nuestra página
if($like_status){
?>
    <!-- Bienvenido a esta página... -->
	<SCRIPT LANGUAGE="JavaScript">
    <!--
    window.location.href = 'tabz.php'
    //-->
    </SCRIPT>
<?php
}else{
?>
<!-- 	<TABLE width="100%">
	<TR>
		<TD align="right"><IMG SRC="http://pda.desa.cl/includes/php/hazclick.jpg" BORDER="0" ALT="Haz click en me gusta"></TD>
	</TR>
	</TABLE>
	<BR><BR><BR><BR> -->

<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"
  codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,42,0"
  id="hazclick2" width="800" height="600">
  <param name="movie" value="hazclick2.swf">
  <param name="bgcolor" value="#FFFFFF">
  <param name="quality" value="high">
  <param name="allowscriptaccess" value="samedomain">
  <embed type="application/x-shockwave-flash"
   pluginspage="http://www.macromedia.com/go/getflashplayer"
   width="800" height="600"
   name="hazclick2" src="hazclick2.swf"
   bgcolor="#FFFFFF" quality="high"
   swLiveConnect="true" allowScriptAccess="samedomain"
  ></embed>
</object>

<?php
}
?>
<!-- <CENTER><FONT SIZE="1" face='Arial' COLOR="#000033">Copyright 2013. Social Advisors. Todos los derechos reservados</FONT></CENTER> -->