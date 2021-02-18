<#
.SYNOPSIS
    NoSync.ps1 - Script para verificar os usuarios que não acessam o AD e o Exchange a um determinado tempo.
.DESCRIPTION 
    checa os todos os usuarios que não acessam a caixa de correio e nem a conta do AD, e adiciona o parametro "noSync" nesses usuarios, a saida desse script é um arquivo de log com todos os usuarios que foram alterados na execuçao desse script.
.OUTPUTS
    Lista com usuarios Inativos que tiveram o parametro "noSync" adicionado a sua conta.
.EXAMPLE
    .\NoSync.ps1 
    adiciona o parametro "NoSync" aos usuarios que nao acessam o AD e o Exchange a um determinado tempo. 
.EXAMPLE
    .\NoSync.ps1 -clear
    adiciona o parametro "NoSync" aos usuarios que nao acessam o AD e o Exchange a mais de 6 meses e verifica se os usuarios que ja tem o parametro "NoSync" acessaram o AD ou o exchange a um determinado tempo.
.NOTES
    Written by: Caio de Amorim Pereira
#>

[CmdletBinding()]
param (
        [Parameter( Mandatory=$false)]
        [switch]$clear
)

#...................................
# Initialize
#...................................

#Add Exchange 2010 snapin if not already loaded in the PowerShell session
if (!(Get-PSSnapin | Where-Object {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
{

	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction STOP
	. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
	Connect-ExchangeServer -auto -AllowClobber

}

#variables 
$lastDate = 30
$days = 180
$now = get-date
$timeLimit = $now.AddDays(-$lastDate)
$inactiveTime = $now.AddDays(-$days)
$lastLogonExchangeUser = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails Usermailbox | where WhenMailboxCreated -le $timeLimit| Get-MailboxStatistics | where LastLogontime -lt $inactiveTime | sort-object DisplayName | Select-Object DisplayName, LastLogontime, LastLogonDate ,TotalItemSize, SamAccountName 
$lastLogonADUser = Get-ADUser -Filter {EmailAddress -notlike "HealthMailbox*" -and EmailAddress -notlike "SystemMailbox*" }  -SearchBase "OU=Usuarios,OU=MD,DC=defesa,DC=net" -properties * | where {$_.LastLogonDate -lt $inactiveTime -and $_.whenCreated -le $timeLimit} | Select-Object -Property Name, LastLogonDate, SamAccountName  | sort-object -property Name 

#objetos vazios
$InactiveUsersADandExchange = @() #caso o usuario não logue nem no Exchange nem no AD, considerado usuario inativo
$InactiveUsersADandExchangeOBJ = @() #objeto de saída 
$clearNoSyncLog = @()
$noSyncUsers = @()
$clearNoSync = @()

#Main Body
foreach($user in $lastLogonExchangeUser)# foreach com os usuarios do exchange que não acessam a caixa a mais de 365 dias 
{ 
    if($user.displayname -in $lastLogonADUser.name ) #if para verificar os usuarios do exchange que não acessam a caixa de correio e o AD a mais de 365 dais
	{ 

            foreach($ADuser in $lastLogonADUser) #foreach para adicionar o paramentro "lastlogondate" nos usuarios 
		    {
                if($ADuser.name -eq $user.DisplayName) #caso o mesmo usuario esteja na lista dos usuarios que nao acessam a caixa do exchange e não acessem o AD, o if adiciona o campo "LastLogonDate" e o "SamAccountName " nos usuarios do exchange
			    {
                    $user.LastLogonDate = $ADuser.lastlogondate
                    $user.SamAccountName = $ADuser.SamAccountName 
                }
            }
        $InactiveUsersADandExchange += $user
    }
}

#transformation of the output of the main body into objects and joining in a single variable
	foreach($userInactive in $InactiveUsersADandExchange)
	{
			$objectHash = @{
				DisplayName = $userInactive.DisplayName
                SamAccountName = $userInactive.SamAccountName 
				LastLogonExchange = $userInactive.LastLogonTime
                LastLogonAD = $userInactive.LastLogonDate
				size = $userInactive.TotalItemsize
			}	
		$userObj = New-Object PSObject -Property $objectHash
		$InactiveUsersADandExchangeOBJ += $userObj
	}

#adiciona o parametro "NoSync"
foreach($UserNoSync in $InactiveUsersADandExchangeOBJ){
    $temp = $UserNoSync.SamAccountName  
    $temp1 = Get-ADUser -Identity "$temp" -Properties *
    if(!$temp1.extensionAttribute15) {   
        Set-ADObject $temp1 -Add @{"extensionAttribute15" = "NoSync"}
        $noSyncUsers +=  Get-ADUser -Identity "$temp" -Properties * | Select-Object name, SamAccountName , extensionAttribute15
    }
}

#output
if($noSyncUsers){
    $dateLog = Get-Date -Format "yyyy-MM-dd_hhmm" 
    $noSyncUsers | ft | Out-File "./Log/noSyncLogs_$dateLog.log" -Encoding utf8
}

#function -clear 
if($clear){

	$clearNoSync = Get-ADUser -filter 'extensionAttribute15 -like "NoSync"' -Properties * | Select-Object name, SamAccountName , extensionAttribute15 
    foreach($UserClear in $clearNoSync){
    <#Logica do IF:
        Caso o $UserClear(Usuario com o parametro "NoSync" ja setado) Não esteja na lista de usuario que não acessam o AD e o Exchange, será removido o parametro "NoSync"
    #>
        if($UserClear.SamAccountName  -notin $InactiveUsersADandExchangeOBJ.SamAccountName ){
           $tempSam = $UserClear.SamAccountName
            Set-ADUser –Identity "$tempSam" -Clear "extensionattribute15"
            $clearNoSyncLog +=  Get-ADUser -Identity "$tempSam" -Properties * | Select-Object name, SamAccountName , extensionAttribute15
        }
    }
    #output do parametro -clear 
	if ($clearNoSyncLog) 
	{
        $dateLog = Get-Date -Format "yyyy-MM-dd_hhmm" 
		$clearNoSyncLog | Out-File "./Log/ClearNoSync_$dateLog.log" -Encoding utf8
	}
}
