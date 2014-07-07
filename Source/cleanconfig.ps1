[string]$filename = gi *.config -Exclude "*vshost*"
[xml]$xml = gc $filename
foreach($n in $xml.configuration.appSettings.add)
{
  switch($n.key)
  {
    "SiteUrl" { $n.value = "http://yoursite.com" }
    "Domain" { $n.value = "" }
	"Username" { $n.value = "Username" }
	"Password" { $n.value = "Password" }
	"FromEmail" { $n.value = "you@you.com" }
	"CcEmail" { $n.value = "you@you.com" }
	"ErrorTo" { $n.value = "you@you.com" }
	"SMTPServer" { $n.value = "smtp.gmail.com" }
	"SMTPPort" { $n.value = "587" }
	"SMTPSSL" { $n.value = "true" }
	"SMTPUserName" { $n.value = "you@gmail.com" }
	"SMTPPassword" { $n.value = "(Not your usual gmail password. Use the Device Password generator in gmail to generate password)" }
	
  } 
}

$xml.Save($filename);