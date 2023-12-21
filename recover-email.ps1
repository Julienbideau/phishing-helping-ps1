# Première requête
$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$session.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
$session.Cookies.Add((New-Object System.Net.Cookie("connect.sid", "s%3AFKbu2Sbf31ic7DuReLizUBWM9admeo6G.H2A0YfMIJ0T%2F3s8QkXyqDItmDf2OcSAM8WHfC1ID0MU", "/", "$urlsite")))
$urlsite = "https://copyurlfrommail.com"
$response = Invoke-WebRequest -UseBasicParsing -Uri "$urlsite/socket.io/?EIO=4&transport=polling&t=OoC3ReX" `
-WebSession $session `
-Headers @{
    "Accept"="*/*"
    "Accept-Encoding"="gzip, deflate, br"
    "Accept-Language"="fr-FR,fr;q=0.9"
    "Referer"="$urlsite/annuler-commande?token=qWLGt0vpWSrzW0UoW"
    "Sec-Fetch-Dest"="empty"
    "Sec-Fetch-Mode"="cors"
    "Sec-Fetch-Site"="same-origin"
    "sec-ch-ua"="`"Not_A Brand`";v=`"8`", `"Chromium`";v=`"120`", `"Google Chrome`";v=`"120`""
    "sec-ch-ua-mobile"="?0"
    "sec-ch-ua-platform"="`"Windows`""
}
# Stocker la réponse JSON dans une variable
$jsonResponse = $response.Content.Substring(1) | ConvertFrom-Json

# Extraire les informations nécessaires
$sessionId = $jsonResponse.sid
$upgrades = $jsonResponse.upgrades

# Deuxième requête avec les informations extraites
$secondResponse = Invoke-WebRequest -UseBasicParsing -Uri "$urlsite/socket.io/?EIO=4&transport=polling&t=OoC3Rh5&sid=$sessionId" `
-Method "POST" `
-WebSession $session `
-Headers @{
    "Accept"="*/*"
    "Accept-Encoding"="gzip, deflate, br"
    "Accept-Language"="fr-FR,fr;q=0.9"
    "Origin"="$urlsite"
    "Referer"="$urlsite/annuler-commande?token=qWLGt0vpWSrzW0UoW"
    "Sec-Fetch-Dest"="empty"
    "Sec-Fetch-Mode"="cors"
    "Sec-Fetch-Site"="same-origin"
    "sec-ch-ua"="`"Not_A Brand`";v=`"8`", `"Chromium`";v=`"120`", `"Google Chrome`";v=`"120`""
    "sec-ch-ua-mobile"="?0"
    "sec-ch-ua-platform"="`"Windows`""
} `
-ContentType "text/plain;charset=UTF-8" `
-Body "40"

# Afficher la deuxième réponse
#$secondResponse.Content

# Troisième requête avec les informations extraites
$thirdResponse = Invoke-WebRequest -UseBasicParsing -Uri "$urlsite/socket.io/?EIO=4&transport=polling&t=OoC3Rh6&sid=$sessionId" `
-WebSession $session `
-Headers @{
    "Accept"="*/*"
    "Accept-Encoding"="gzip, deflate, br"
    "Accept-Language"="fr-FR,fr;q=0.9"
    "Referer"="$urlsite/annuler-commande?token=qWLGt0vpWSrzW0UoW"
    "Sec-Fetch-Dest"="empty"
    "Sec-Fetch-Mode"="cors"
    "Sec-Fetch-Site"="same-origin"
    "sec-ch-ua"="`"Not_A Brand`";v=`"8`", `"Chromium`";v=`"120`", `"Google Chrome`";v=`"120`""
    "sec-ch-ua-mobile"="?0"
    "sec-ch-ua-platform"="`"Windows`""
}

$thirdResponse.Content.Substring(0,500)	

# Utiliser une expression régulière pour extraire le texte après la chaîne ["historicalData",
$pattern = '\["historicalData",(.*)'

# Faire correspondre la chaîne avec le motif
$match = [regex]::Match($thirdResponse.Content, $pattern)

# Vérifier si la correspondance est trouvée
if ($match.Success) {
    # Extraire le texte après ["historicalData",
    $resultat = $match.Groups[1].Value
    $jsonArray = $resultat.Substring(0, $resultat.Length - 1) | ConvertFrom-Json
	$lastStepPeople = $jsonArray | Where-Object { $_.pageName -like '*VBV*' }
	$groupedResults = $lastStepPeople | Group-Object userEmail | ForEach-Object {
		$latestDate = $_.Group | Measure-Object date -Maximum | Select-Object -ExpandProperty Maximum
		[PSCustomObject]@{
			userEmail = $_.Name
			latestDate = $latestDate
		}
	}

	# Afficher les résultats groupés
	$groupedResults
		# Accéder aux éléments du tableau
	foreach ($item in $lastStepPeople) {
		Write-Host "ID: $($item.id)"
		Write-Host "User Email: $($item.userEmail)"
		Write-Host "Page Name: $($item.pageName)"
		Write-Host "Date: $($item.date)"
	}
}
