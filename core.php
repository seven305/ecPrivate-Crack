<?php
$login = $_POST['username'];
$pass = $_POST['pass'];

$url = "https://my.screenname.aol.com/_cqr/login/login.psp";
$user_agent = "Mozilla/5.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; .NET CLR 1.1.4322; .NET CLR 2.0.50727)";
$postvars = "screenname=".$login."&password=".$pass."&sitedomain=pictures.aol.com&lang=es&locale=us&authLev=1&mcState=initialized&siteState=OrigUrl%3Dhttp%3A%2F%2Ffotoslatino.aol.com%3A80%2Fap%2FcreateAlbum.do";

$cookie = "cookies/test";
@unlink($cookie);

$ch = curl_init($url);
curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, 1);
curl_setopt($ch, CURLOPT_USERAGENT, $user_agent);
curl_setopt($ch, CURLOPT_HEADER, 1);
curl_setopt($ch, CURLOPT_POST, 1);
curl_setopt($ch, CURLOPT_POSTFIELDS, $postvars);
curl_setopt($ch, CURLOPT_COOKIEJAR, $cookie);
curl_setopt($ch, CURLOPT_FOLLOWLOCATION, 1);
curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, TRUE);
$content = curl_exec ($ch);

//file_get_contents("http://sns.webmail.aol.com");
//print_r(curl_getinfo($ch)); 
//echo "\n\ncURL error number:" .curl_errno($ch); 
//echo "\n\ncURL error:" . curl_error($ch);  
curl_close ($ch);
unset($ch);
echo "<PRE>".htmlentities($content); 

?>