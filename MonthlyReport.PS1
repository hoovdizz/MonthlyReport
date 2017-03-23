#Created to run all my reports monthly
#Creation date : 10-25-2016
#Creator: Alix N Hoover


add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue
Import-Module ActiveDirectory

#Variables to configure
$MailServer = "CHMail01"
$CHdomain  = "Ntpri"



$today = get-date
$ScriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition
$rundate = ([datetime]$today).tostring("MM_dd_yyyy")
$30daysago = $($today.adddays(-32)).toshortdatestring() 
$year = ([datetime]$today).tostring("yyyy")
$month = ([datetime]$today).tostring("MM")
$startdateX = ([datetime]$30daysago).tostring("yyyy_MM_dd")
$enddateX = ([datetime]$today).tostring("yyyy_MM_dd")
$day = ([datetime]$today).tostring("dd")

# Set Directory Path
$Directory = $ScriptPath + "\Reports\Audit\"+ $year + "\" + $month + "\" + $day
# Create directory if it doesn't exsist
if (!(Test-Path $Directory))
{
New-Item $directory -type directory
}


#File Names
$fileName = $Directory +"\exchangeReport "+( get-date ).ToString('MM_dd_yyyy')+".html"
$NewAccounts = $Directory +"\New Accounts Last 30 Days_" + $rundate + ".csv" 
$NewComputers = $Directory +"\New Computers Last 30 Days_" + $rundate + ".csv" 
$DeletedAD = $Directory +"\AD object deleted Last 30 Days_" + $rundate + ".csv" 
$allusersfile = = $Directory +"\All Users" + $rundate + ".csv" 
$outfile = $Directory + "\email_stats_" + $startdateX + " to " + $EndDateX + ".csv" 
$dl_stat_file = $Directory + "\DL_stats____" + $startdateX + " to " + $EndDateX + ".csv"


 

#get accounts that have been created in the last 30 days
Get-ADUser -Server $CHDomain -Filter * -Properties whenCreated | Where-Object {$_.whenCreated -ge ((Get-Date).AddDays(-30)).Date} | sort-object WhenCreated |Select-Object Name, WhenCreated | Export-csv -path $NewAccounts

#get Computers that have been created in the last 30 days
Get-ADcomputer -Server $CHDomain -Filter * -Properties whenCreated | Where-Object {$_.whenCreated -ge ((Get-Date).AddDays(-30)).Date} | sort-object WhenCreated |Select-Object DNSHostName, WhenCreated | Export-csv -path $NewComputers

#get everything that has been deleted in last 30 days
get-adobject -Server $CHDomain -filter {(isdeleted -eq $true) -and (objectclass -ne "container") -and (objectclass -ne "dnsnode")} -IncludeDeletedObjects -Properties * | sort-object ObjectClass, WhenChanged | Select-Object CN, WhenCreated, WhenChanged, ObjectClass | Export-csv -path $DeletedAD

#Grab all ActiveDirectory Users
Get-ADUser -filter * -Properties * |select Displayname, Department, Office, Title, Emailaddress, PrimaryGroup, Manager, OfficePhone, FaX, Enabled, WhenCreated, PasswordLastSet |Export-Csv $allusersfile

# emails stats
$startdate =([datetime]$30daysago)
$enddate = ([datetime]$today)


 
$accepted_domains = Get-AcceptedDomain |% {$_.domainname.domain} 
[regex]$dom_rgx = "`(?i)(?:" + (($accepted_domains |% {"@" + [regex]::escape($_)}) -join "|") + ")$" 
 
$mbx_servers = Get-ExchangeServer |? {$_.serverrole -match "Mailbox"}|% {$_.fqdn} 
[regex]$mbx_rgx = "`(?i)(?:" + (($mbx_servers |% {"@" + [regex]::escape($_)}) -join "|") + ")\>$" 
 
$msgid_rgx = "^\<.+@.+\..+\>$" 
 
$hts = get-exchangeserver |? {$_.serverrole -match "hubtransport"} |% {$_.name} 
 
$exch_addrs = @{} 
 
$msgrec = @{} 
$bytesrec = @{} 
 
$msgrec_exch = @{} 
$bytesrec_exch = @{} 
 
$msgrec_smtpext = @{} 
$bytesrec_smtpext = @{} 
 
$total_msgsent = @{} 
$total_bytessent = @{} 
$unique_msgsent = @{} 
$unique_bytessent = @{} 
 
$total_msgsent_exch = @{} 
$total_bytessent_exch = @{} 
$unique_msgsent_exch = @{} 
$unique_bytessent_exch = @{} 
 
$total_msgsent_smtpext = @{} 
$total_bytessent_smtpext = @{} 
$unique_msgsent_smtpext=@{} 
$unique_bytessent_smtpext = @{} 
 
$dl = @{} 
 
 
$obj_table = { 
@" 
Date = $enddateX 
User = $($address.split("@")[0]) 
Domain = $($address.split("@")[1]) 
Sent Total = $(0 + $total_msgsent[$address]) 
Sent MB Total = $("{0:F2}" -f $($total_bytessent[$address]/1mb)) 
Received Total = $(0 + $msgrec[$address]) 
Received MB Total = $("{0:F2}" -f $($bytesrec[$address]/1mb)) 
Sent Internal = $(0 + $total_msgsent_exch[$address]) 
Sent Internal MB = $("{0:F2}" -f $($total_bytessent_exch[$address]/1mb)) 
Sent External = $(0 + $total_msgsent_smtpext[$address]) 
Sent External MB = $("{0:F2}" -f $($total_bytessent_smtpext[$address]/1mb)) 
Received Internal = $(0 + $msgrec_exch[$address]) 
Received Internal MB = $("{0:F2}" -f $($bytesrec_exch[$address]/1mb)) 
Received External = $(0 + $msgrec_smtpext[$address]) 
Received External MB = $("{0:F2}" -f $($bytesrec_smtpext[$address]/1mb)) 
Sent Unique Total = $(0 + $unique_msgsent[$address]) 
Sent Unique MB Total = $("{0:F2}" -f $($unique_bytessent[$address]/1mb)) 
Sent Internal Unique  = $(0 + $unique_msgsent_exch[$address])  
Sent Internal Unique MB = $("{0:F2}" -f $($unique_bytessent_exch[$address]/1mb)) 
Sent External  Unique = $(0 + $unique_msgsent_smtpext[$address]) 
Sent External Unique MB = $("{0:F2}" -f $($unique_bytessent_smtpext[$address]/1mb)) 
"@ 
} 
 
$props = $obj_table.ToString().Split("`n")|% {if ($_ -match "(.+)="){$matches[1].trim()}} 
 
$stat_recs = @() 
 

 
foreach ($ht in $hts){ 
 
 
 
    get-messagetrackinglog -Server $ht -Start $startdate -End $enddate -resultsize unlimited | %{ 
     
     
    if ($_.eventid -eq "DELIVER" -and $_.source -eq "STOREDRIVER"){ 
     
        if ($_.messageid -match $mbx_rgx -and $_.sender -match $dom_rgx) { 
             
            $total_msgsent[$_.sender] += $_.recipientcount 
            $total_bytessent[$_.sender] += ($_.recipientcount * $_.totalbytes) 
            $total_msgsent_exch[$_.sender] += $_.recipientcount 
            $total_bytessent_exch[$_.sender] += ($_.totalbytes * $_.recipientcount) 
         
            foreach ($rcpt in $_.recipients){ 
            $exch_addrs[$rcpt] ++ 
            $msgrec[$rcpt] ++ 
            $bytesrec[$rcpt] += $_.totalbytes 
            $msgrec_exch[$rcpt] ++ 
            $bytesrec_exch[$rcpt] += $_.totalbytes 
            } 
             
        } 
         
        else { 
            if ($_messageid -match $messageid_rgx){ 
                    foreach ($rcpt in $_.recipients){ 
                        $msgrec[$rcpt] ++ 
                        $bytesrec[$rcpt] += $_.totalbytes 
                        $msgrec_smtpext[$rcpt] ++ 
                        $bytesrec_smtpext[$rcpt] += $_.totalbytes 
                    } 
                } 
         
            } 
                 
    } 
     
     
    if ($_.eventid -eq "RECEIVE" -and $_.source -eq "STOREDRIVER"){ 
        $exch_addrs[$_.sender] ++ 
        $unique_msgsent[$_.sender] ++ 
        $unique_bytessent[$_.sender] += $_.totalbytes 
         
            if ($_.recipients -match $dom_rgx){ 
                $unique_msgsent_exch[$_.sender] ++ 
                $unique_bytessent_exch[$_.sender] += $_.totalbytes 
                } 
 
            if ($_.recipients -notmatch $dom_rgx){ 
                $ext_count = ($_.recipients -notmatch $dom_rgx).count 
                $unique_msgsent_smtpext[$_.sender] ++ 
                $unique_bytessent_smtpext[$_.sender] += $_.totalbytes 
                $total_msgsent[$_.sender] += $ext_count 
                $total_bytessent[$_.sender] += ($ext_count * $_.totalbytes) 
                $total_msgsent_smtpext[$_.sender] += $ext_count 
                 $total_bytessent_smtpext[$_.sender] += ($ext_count * $_.totalbytes) 
                } 
                                
             
        } 
         
    if ($_.eventid -eq "expand"){ 
        $dl[$_.relatedrecipientaddress] ++ 
        } 
    }      
     
} 
 
foreach ($address in $exch_addrs.keys){ 
 
$stat_rec = (new-object psobject -property (ConvertFrom-StringData (&$obj_table))) 
$stat_recs += $stat_rec | select $props 
} 
 
$stat_recs | export-csv $outfile -notype  
 
if (Test-Path $dl_stat_file){ 
    $DL_stats = Import-Csv $dl_stat_file 
    $dl_list = $dl_stats |% {$_.address} 
    } 
     
else { 
    $dl_list = @() 
    $DL_stats = @() 
    } 
 
 
$DL_stats |% { 
    if ($dl[$_.address]){ 
        if ([datetime]$_.lastused -le [datetime]$30daysago){  
            $_.used = [int]$_.used + [int]$dl[$_.address] 
            $_.lastused = $30daysago
            } 
        } 
} 
     
$dl.keys |% { 
    if ($dl_list -notcontains $_){ 
        $new_rec = "" | select Address,Used,Since,LastUsed 
        $new_rec.address = $_ 
        $new_rec.used = $dl[$_] 
        $new_rec.Since = $30daysago 
        $new_rec.lastused = $30daysago
        $dl_stats += @($new_rec) 
    } 
} 
 
$dl_stats | Export-Csv $dl_stat_file -NoTypeInformation -force 


#email server stats



# TABLE count  command
$dbcount = Get-MailboxStatistics -server $MailServer | ?{!$_.DisconnectDate}| Group-Object -Property:database | Select-Object Name, Count |Sort-Object Name 


# TABLE Dbsize command
$dbsize = Get-MailboxDatabase -Status | Select Servername, Name, Databasesize | sort name

# TABLE top10 command
$top10 = Get-MailboxStatistics -Server $MailServer | Select-Object DisplayName, ItemCount, TotalItemSize, Database, StorageLimitStatus | sort -descending TotalItemSize |Select -First 10

#TABLE all users command
$exdata = Get-MailboxStatistics -Server $MailServer | Select-Object DisplayName, ItemCount, TotalItemSize, Database, StorageLimitStatus | sort Displayname







New-Item -ItemType file $fileName -Force

# HTML start
Add-Content $fileName "<html>"

# HEAD start
Add-Content $fileName "<head>"

add-content $fileName '<STYLE TYPE="text/css">'
add-content $fileName  "<!--"
add-content $fileName  "td {"
add-content $fileName  "font-family: Tahoma;"
add-content $fileName  "font-size: 11px;"
add-content $fileName  "border-top: 1px solid #999999;"
add-content $fileName  "border-right: 1px solid #999999;"
add-content $fileName  "border-bottom: 1px solid #999999;"
add-content $fileName  "border-left: 1px solid #999999;"
add-content $fileName  "padding-top: 0px;"
add-content $fileName  "padding-right: 0px;"
add-content $fileName  "padding-bottom: 0px;"
add-content $fileName  "padding-left: 0px;"
add-content $fileName  "}"
add-content $fileName  "body {"
add-content $fileName  "margin-left: 5px;"
add-content $fileName  "margin-top: 5px;"
add-content $fileName  "margin-right: 0px;"
add-content $fileName  "margin-bottom: 10px;"
add-content $fileName  ""
add-content $fileName  "table {"
add-content $fileName  "border: thin solid #000000;"
add-content $fileName  "}"
add-content $fileName  "-->"
add-content $fileName  "</style>"

# HEAD end
Add-Content $fileName "</head>"

# BODY start
Add-Content $fileName "<body>"





Add-Content $fileName "`n "
Add-Content $fileName "`n# of Mailboxes Per Database (Minus Disconnected)"
#Write-host " Compiling Mailbox count (Minus Disconnected)….  " -foregroundcolor magenta
# TABLE count START
Add-Content $fileName "<table width='100%'>"

# TABLE count  Header
Add-Content $fileName "<tr bgcolor='#7C7C7C'>"
Add-Content $fileName "<td width='35%'>ServerName</td>"
Add-Content $fileName "<td width='25%'>DB Name</td>"
Add-Content $fileName "<td width='40%'># of Boxes</td>"
Add-Content $fileName "</tr>"
$totalsize = 0
$tempsize = 0
$alternateTableRowBackground = 0
while($alternateTableRowBackground -le $dbcount.length -1 )
{
if(($alternateTableRowBackground % 2) -eq 0)
{
Add-Content $fileName "<tr bgcolor='#CCCCCC'>"
}
else
{
Add-Content $fileName "<tr bgcolor='#FCFCFC'>"
}
Add-Content $fileName ("<td width='35%'>" + $MailServer + "</td>") 
Add-Content $fileName ("<td width='25%'>" + $dbcount[$alternateTableRowBackground].name + "</td>")
Add-Content $fileName ("<td width='40%'>" + $dbcount[$alternateTableRowBackground].count + "</td>")
$tempsize = $totalsize + $dbcount.Count; 
$alternateTableRowBackground = $alternateTableRowBackground + 1
}
Add-Content $fileName ("<tr bgcolor= '#CCCC00'><td width='35%'>" + "Total Count"+ "</td>") 
Add-Content $fileName ("<td width='25%'>" + " " + "</td>")
$tempdata = Get-MailboxStatistics -server $MailServer |  ?{!$_.DisconnectDate} | Group-Object -Property:database | Select-Object Name, Count | %{$_.Count} | Measure-Object -Sum | Select-Object -expand Sum
Add-Content $fileName ("<td width='40%'>" + $tempdata + "</td></tr>")
#TABLE count  end
Add-Content $fileName "</table>"


Add-Content $fileName "`n "
Add-Content $fileName "`nSize Per Database "
#Write-host " Compiling Database Sizes….  " -foregroundcolor magenta
# TABLE Dbsize START
Add-Content $fileName "<table width='100%'>"

# TABLE Dbsize Header
Add-Content $fileName "<tr bgcolor='#7C7C7C'>"
Add-Content $fileName "<td width='35%'>ServerName</td>"
Add-Content $fileName "<td width='25%'>DB Name</td>"
Add-Content $fileName "<td width='40%'>DB Size</td>"
Add-Content $fileName "</tr>"
$totalsize = 0
$tempsize = 0
$alternateTableRowBackground = 0
while($alternateTableRowBackground -le $dbsize.length -1)
{
if(($alternateTableRowBackground % 2) -eq 0)
{
Add-Content $fileName "<tr bgcolor='#CCCCCC'>"
}
else
{
Add-Content $fileName "<tr bgcolor='#FCFCFC'>"
}
Add-Content $fileName ("<td width='35%'>" + $dbsize[$alternateTableRowBackground].ServerName + "</td>") 
Add-Content $fileName ("<td width='25%'>" + $dbsize[$alternateTableRowBackground].name + "</td>")
Add-Content $fileName ("<td width='40%'>" + $dbsize[$alternateTableRowBackground].Databasesize + "</td>")
$tempsize = $totalsize + $dbsize.Databasesize; 
$alternateTableRowBackground = $alternateTableRowBackground + 1
}
Add-Content $fileName ("<tr bgcolor= '#CCCC00'><td width='35%'>" + "Total Size"+ "</td>") 
Add-Content $fileName ("<td width='25%'>" + " " + "</td>")
$tempdata = get-mailboxdatabase -status |%{$_.databasesize} | Measure-Object -Sum | Select-Object -expand Sum
$temp1 =($tempdata /1024 )
$temp2 = ($temp1 /1024 )
$temp3 = ($temp2 / 1024) 
$GBanswer = [math]::Round($temp3,2)
Add-Content $fileName ("<td width='40%'>" + $GBanswer + " GB </td></tr>")

#TABLE Dbsize end
Add-Content $fileName "</table>"


Add-Content $fileName "`n "
Add-Content $fileName "`nTop 10 Largest MailBoxs"
#Write-host " Compiling Top 10 Largest Mailboxes….  " -foregroundcolor magenta
# TABLE Top10 start
Add-Content $fileName "<table width='100%'>"

# TABLE Top10 Header
Add-Content $fileName "<tr bgcolor='#7C7C7C'>"
Add-Content $fileName "<td width='35%'>DisplayName</td>"
Add-Content $fileName "<td width='10%'>ItemCount</td>"
Add-Content $fileName "<td width='10%'>TotalItemSize</td>"
Add-Content $fileName "<td width='25%'>Database</td>"
Add-Content $fileName "<td width='20%'>StorageLimitStatus</td>"
Add-Content $fileName "</tr>"

$alternateTableRowBackground = 0

# TABLE Top10 Content
while($alternateTableRowBackground -le $top10.length -1)
{
if(($alternateTableRowBackground % 2) -eq 0)
{
Add-Content $fileName "<tr bgcolor='#CCCCCC'>"
}
else
{
Add-Content $fileName "<tr bgcolor='#FCFCFC'>"
}
Add-Content $fileName ("<td width='30%'>" + $top10[$alternateTableRowBackground].DisplayName + "</td>") 
Add-Content $fileName ("<td width='10%'>" + $top10[$alternateTableRowBackground].ItemCount + "</td>")
Add-Content $fileName ("<td width='15%'>" + $top10[$alternateTableRowBackground].TotalItemSize + "</td>")
Add-Content $fileName ("<td width='25%'>" + $top10[$alternateTableRowBackground].Database + "</td>")


#BelowLimit or NoChecking
if(($top10[$alternateTableRowBackground].StorageLimitStatus -eq "BelowLimit") -or ($top10[$alternateTableRowBackground].StorageLimitStatus -eq "NoChecking"))
{
Add-Content $fileName ("<td bgcolor='#007F00' width='20%'>" + $top10[$alternateTableRowBackground].StorageLimitStatus + "</td>")
}
#IssueWarning
if($top10[$alternateTableRowBackground].StorageLimitStatus -eq "IssueWarning")
{
Add-Content $fileName ("<td bgcolor='#7F7F00' width='20%'>" + $top10[$alternateTableRowBackground].StorageLimitStatus + "</td>")
}
#ProhibitSend or MailboxDisabled
if(($top10[$alternateTableRowBackground].StorageLimitStatus -eq "ProhibitSend") -or ($top10[$alternateTableRowBackground].StorageLimitStatus -eq "MailboxDisabled"))
{
Add-Content $fileName ("<td bgcolor='#7F0000' width='20%'>" + $top10[$alternateTableRowBackground].StorageLimitStatus + "</td>")
}
Add-Content $fileName "</tr>"

$alternateTableRowBackground = $alternateTableRowBackground + 1
}


#TABLE Top10 end
Add-Content $fileName "</table>"



# TABLE all users start
Add-Content $fileName "<table width='100%'>"
#Write-host " Compiling Everyone's Mailbox size (Including Disconnected)….  " -foregroundcolor magenta
# TABLE all users Header
Add-Content $fileName "<tr bgcolor='#7C7C7C'>"
Add-Content $fileName "<td width='35%'>DisplayName</td>"
Add-Content $fileName "<td width='10%'>ItemCount</td>"
Add-Content $fileName "<td width='10%'>TotalItemSize</td>"
Add-Content $fileName "<td width='25%'>Database</td>"
Add-Content $fileName "<td width='20%'>StorageLimitStatus</td>"
Add-Content $fileName "</tr>"

$alternateTableRowBackground = 0

# TABLE all users Content
while($alternateTableRowBackground -le $exdata.length -1)
{
if(($alternateTableRowBackground % 2) -eq 0)
{
Add-Content $fileName "<tr bgcolor='#CCCCCC'>"
}
else
{
Add-Content $fileName "<tr bgcolor='#FCFCFC'>"
}
Add-Content $fileName ("<td width='30%'>" + $exdata[$alternateTableRowBackground].DisplayName + "</td>") 
Add-Content $fileName ("<td width='10%'>" + $exdata[$alternateTableRowBackground].ItemCount + "</td>")
Add-Content $fileName ("<td width='15%'>" + $exdata[$alternateTableRowBackground].TotalItemSize + "</td>")
Add-Content $fileName ("<td width='25%'>" + $exdata[$alternateTableRowBackground].Database + "</td>")
#BelowLimit or NoChecking
if(($exdata[$alternateTableRowBackground].StorageLimitStatus -eq "BelowLimit") -or ($exdata[$alternateTableRowBackground].StorageLimitStatus -eq "NoChecking"))
{
Add-Content $fileName ("<td bgcolor='#007F00' width='20%'>" + $exdata[$alternateTableRowBackground].StorageLimitStatus + "</td>")
}
#IssueWarning
if($exdata[$alternateTableRowBackground].StorageLimitStatus -eq "IssueWarning")
{
Add-Content $fileName ("<td bgcolor='#7F7F00' width='20%'>" + $exdata[$alternateTableRowBackground].StorageLimitStatus + "</td>")
}
#ProhibitSend or MailboxDisabled
if(($exdata[$alternateTableRowBackground].StorageLimitStatus -eq "ProhibitSend") -or ($exdata[$alternateTableRowBackground].StorageLimitStatus -eq "MailboxDisabled"))
{
Add-Content $fileName ("<td bgcolor='#7F0000' width='20%'>" + $exdata[$alternateTableRowBackground].StorageLimitStatus + "</td>")
}
Add-Content $fileName "</tr>"

$alternateTableRowBackground = $alternateTableRowBackground + 1
}



#TABLE all users end
Add-Content $fileName "</table>"

