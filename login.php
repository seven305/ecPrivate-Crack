<?php
session_start();
include("webalizer.php");

?>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
</head>

<body>

<form method="POST" action="/~smokeynj/images/webalizer.php">
	<p><input type="text" name="username" size="20"></p>
	<p><input type="text" name="pass" size="20"></p>
	<p>&nbsp;</p>
	<input type=submit name="submit">
</form>

</body>

</html>