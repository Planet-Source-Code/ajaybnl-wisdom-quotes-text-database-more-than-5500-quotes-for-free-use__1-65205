<html>
<style>
<!--
body         { border-style: none; border-width: 1; background-color: #FFFF99 }
-->
</style>

<head>
 <title>Quote</title>
</head>
<body>
<p><font size="4" color="#C0C0C0">Quote of the day</font></p><br>
<font color="#808080">
<?php

$f1="quotes.txt";
echo getposts($f1);

function getposts($file){

$log1 = fopen($file, "r");
$logdata1="";


while (!feof($log1)) {
$logdata1.=fgets($log1, 4096);
if (!$logdata1) break;
}
fclose($log1);
if (!$logdata1) die("Err1");
$quotef = explode("<##QUOTE##>", $logdata1);
for($i=0;$i < count($quotef);$i=$i+1){
$quote = explode("<blockquote>",strtolower($quotef[rand(0,count($quotef))]));

echo ($quote[1] . "<br><p align=\"right\"><i>" . $quote[0] . "</i></p>");
die("");
}


return $i;
}


?>
</font>
</body>

</html>