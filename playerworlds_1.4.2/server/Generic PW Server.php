<?php
?>
<html>
<title>Generic PW Server</title>
<body background= "../backgrounds/bolt.jpg"link = white vlink=yellow > 
<br><font size=4 color = #C7001D >
<h1 align = center>Generic PW Server</h1>
</font>
<br><font size=4 color = #C7001D >Server Name:</font><font size=4 color= #FFFFFF> Generic PW Server</font>
<br><font size=4 color = #C7001D >IP Address :</font><font size=4 color= #FFFFFF> 66.25.100.240</font>
<br><font size=4 color = #C7001D >Port Number:</font><font size=4 color= #FFFFFF> 7234</font>
<br><font size=4 color = #C7001D >Website : <a href =
http://www.playerworlds.com.com>
http://www.playerworlds.com.com</a>
<br><font size=4 color = #C7001D >Status : </font>
<?
$ip = "66.25.100.240";
$port = "7234";
if (! $sock = fsockopen($ip, $port, $num, $error, 5))
 echo '<B><FONT COLOR=red>OFFLINE</b></FONT>';
else{
 echo '<B><FONT COLOR=lime>ONLINE</b></FONT>';
fclose($sock);
}
?>
<br><font size=4 color = #C7001D >Server Owner :</font><font size=4 color= #FFFFFF> Kael</font>
<br><font size=4 color = #C7001D >Email: </font><font size=4 color= #FFFFFF> kael@playerworlds.com<hr></font>
<p><h3><b><u><br><font size=4 color = #C7001D >Description :</u></b></h3></font><font size=4 color= #FFFFFF> 
<br><B></font><font color = #c7001D size = 4> This is the Official Player Worlds &copy Beta testing server</B>.</font><font color = #FFFFFF><br> The client for this server is not currently
<br>available to everyone, however you can access the server via the current JO client.
<br>Any new updates will be posted on the Jerrath website, so check the news section frequently for updates.
<br>Thanks,
<br>~Kael~</font>
</font></body></html>
