

# notice Darvis Analyse Poste V1.00
# notice Darvis Analyse Poste V1.00 
# .\AnalyseSysteme.ps1
# Paramètre ligne de commande
# Paramètre ligne de commande 
# .\AnalyseSysteme.ps1 -NoEmail
# .\AnalyseSysteme.ps1 -KeepPreviousReport
# .\AnalyseSysteme.ps1 -KeepPreviousReport 

# Création du fichier mot de passe
# "votre_mot_de_passe" | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Set-Content -Path "Param\password.txt"
# Création du fichier mot de passe 
#"votre_mot_de_passe" | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Set-Content -Path "Param\password.txt"

# Configuration
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$ConfigFile = Join-Path $ScriptPath "Param\config.json"
$PasswordFile = Join-Path $ScriptPath "Param\password.txt"
$ReportFolder = Join-Path $ScriptPath "rapport"
$Today = Get-Date -Format "yyyyMMdd"
$ComputerName = $env:COMPUTERNAME
$ReportFile = Join-Path $ReportFolder "${Today}_${ComputerName}.html"

# Charger les paramètres depuis le fichier de configuration
if (-Not (Test-Path $ConfigFile)) {
    Write-Error "Le fichier de configuration est manquant."
    exit
}
$Config = Get-Content $ConfigFile | ConvertFrom-Json

# Déchiffrer le mot de passe
$SecurePassword = Get-Content $PasswordFile | ConvertTo-SecureString

# Options de lancement du script
param (
    [switch]$NoEmail,
    [switch]$KeepPreviousReport
)

# Suppression du rapport de la veille si l'option $KeepPreviousReport n'est pas activée
if (-Not $KeepPreviousReport) {
    Get-ChildItem -Path $ReportFolder -Filter "*.html" | Where-Object { $_.Name -ne "${Today}_${ComputerName}.html" } | Remove-Item -Force
}

# Fonction pour analyser l'espace disque
function Get-DiskSpaceReport {
    Get-PSDrive -PSProvider FileSystem | Select-Object Name, @{Name="Free(%)"; Expression={"{0:P2}" -f ($_.Free/$_.Used)}}, @{Name="Free(GB)"; Expression={[math]::Round($_.Free/1GB, 2)}}
}

# Vérification des erreurs Windows des 24 dernières heures avec exclusions
function Get-WindowsErrors {
    $ExcludedEvents = $Config.ExcludedEventIDs
    Get-WinEvent -FilterHashtable @{
        LogName= @('Application', 'System', 'Security');
        Level=1,2;
        StartTime=(Get-Date).AddHours(-24)
    } 
    Where-Object { -not ($ExcludedEvents -contains $_.Id) } |
    Select-Object TimeCreated, Id, LevelDisplayName, Message -First 5
}

# Générer le rapport HTML
function Generate-Report {
    $DiskSpace = Get-DiskSpaceReport
    $WindowsErrors = Get-WindowsErrors

    # Alerte si l'espace disque est inférieur à 15%
    $LowDiskAlert = $DiskSpace | Where-Object { [double]($_.'Free(%)' -replace '%', '') -lt 15 }

    $HtmlReport = @"
    <html>
    <head><title>Rapport d'Analyse - $ComputerName</title></head>
    <body>
    <h1>Rapport d'Analyse du Système</h1>
    <h2>Espace Disque</h2>
    <table border="1">
        <tr><th>Lecteur</th><th>Libres (%)</th><th>Libres (GB)</th></tr>
"@

    foreach ($Disk in $DiskSpace) {
        $HtmlReport += "<tr><td>$($Disk.Name)</td><td>$($Disk.'Free(%)')</td><td>$($Disk.'Free(GB)')</td></tr>"

    }

    $HtmlReport += "</table>"

    if ($LowDiskAlert) {
        $HtmlReport += "<p style='color:red'><strong>Alerte :</strong> Un ou plusieurs lecteurs ont moins de 15% d'espace libre.</p>"
    }

    $HtmlReport += "<h2>Erreurs Windows (24h)</h2><table border='1'><tr><th>Date</th><th>ID</th><th>Niveau</th><th>Message</th></tr>"

    foreach ($Error in $WindowsErrors) {
        $HtmlReport += "<tr><td>$($Error.TimeCreated)</td><td>$($Error.Id)</td><td>$($Error.LevelDisplayName)</td><td>$($Error.Message)</td></tr>"
    }

    $HtmlReport += "</table></body></html>"

    # Sauvegarder le rapport
    $HtmlReport | Out-File -FilePath $ReportFile -Encoding UTF8
}

# Fonction pour envoyer l'e-mail
function Send-Email {
    $EmailSettings = $Config.EmailSettings
    $EmailMessage = @{
        To = $EmailSettings.To
        From = $EmailSettings.From
        Subject = "Rapport d'Analyse - $ComputerName"
        Body = "Veuillez trouver en pièce jointe le rapport d'analyse."
        SmtpServer = $EmailSettings.SmtpServer
        Port = $EmailSettings.Port
        UseSsl = $true
        Credential = New-Object System.Management.Automation.PSCredential ($EmailSettings.Username, $SecurePassword)
        Attachments = $ReportFile
    }

    Send-MailMessage @EmailMessage
}

# Exécution du script
Generate-Report

if (-Not $NoEmail) {
    Send-Email
    Write-Output "Le rapport a été envoyé avec succès."
} else {
    Write-Output "Le rapport a été généré mais non envoyé."
}


