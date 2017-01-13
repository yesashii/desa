<?php require_once 'appinclude.php';?>
<h1>
Mi primer canvas!</p>
Hola <fb:name uid="<?=$user;?>" useyou="false"/></p>
Tus amigos son:</p>
<table>
<?php
$i = 1;
foreach ($facebook->api_client->friends_get() as $friend_id) {
    if ($i == 1){
     echo "<tr>";
   }
   echo "<td>" . "<fb:profile-pic uid='" . $friend_id . "'/>" . "</td>";
   echo "<td>" . "<fb:name uid='" . $friend_id . "'/></br>" . "</td>";
   if ($i == 4) {
     $i = 0;
     echo "</tr>";
   }
   $i++;
 }
 ?>
 </table>
 </h1>