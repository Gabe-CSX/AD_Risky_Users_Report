$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$Error.Clear()
Write-Progress "HTML setup"

$head = @"
<meta charset='UTF-8'>
<meta name='author' content='Gabe-CSX on github'>
<meta name='viewport' content='width=device-width, initial-scale=1.0'>
<title>Risky Users</title>
<!-- Boxicons CSS -->
<link href='https://unpkg.com/boxicons@2.1.1/css/boxicons.min.css' rel='stylesheet'>
<style>
* {
    padding: 0;
    margin: 0;
}

html {
  font-family: Arial, Helvetica, sans-serif;
  font-size: 18px;
  width: 100%;
  margin: 0 auto;
}

.user {
  padding: 10px;
}

h2 {
  color: rgb(13, 218, 207);
  font-weight: bolder;
}

table {
  margin: 0 auto;
  width: 100%;
  border-spacing: 0;
}

td {
  padding: 10px;
}

th {
  color: white;
  background-color: rgb(13, 218, 207);
  padding: 0.125em;
}

table th + th {
  border-left: 1px solid white;
}

#postcontent {
  padding: 10px;
  list-style-type: circle;
  list-style-position: inside;
}

#postcontent li {
    padding: 2rem;
}

.nav {
    position: fixed;
    width: 100%;
    box-shadow: 0 0 10px lightgray;
    background-color: white;
}

a {
    text-decoration: none;
    color: black;
}

.nav_list {
    display: flex;
    justify-content: flex-end;
    align-items: center;
    gap: 2rem;
    margin: 0 10px;
}

.nav_listlogo {
    list-style: none;
    margin-right: auto;
    margin-left: 0;
    cursor: pointer;
}

.nav_listlogo i {
    overflow: hidden;
    transition: 0.5s ease-in-out;
    color: white;
    background-color: black;
    padding: .5rem 2rem;
    border-radius: 5px;
    text-align: right;
}

.nav_listlogo i:hover {
    background-color: rgb(13, 218, 207);
}

.nav_listitem {
    list-style: none;
    font-weight: bold;
    position: relative;
    padding: 1.5rem 1rem;
    cursor: pointer;
    transition: background-color 200ms ease-in-out;
}

.nav_listitem::after{
    content: '';
    width: 0;
    height: 0.3rem;
    border-radius: 5px;
    position: absolute;
    left: 1rem;
    bottom: 0.8rem;
    background-color: rgb(13, 218, 207);
    transform: width 200ms ease-in;
}

.nav_listitem:hover::after {
    width: 80%;
}

.nav_listitem:hover .nav_listitemdrop{
    opacity: 1;
    visibility: visible;
}

.nav_listitemdrop {
    position: absolute;
    top: 4rem;
    left: -1rem;
    box-shadow: 0 0 10px lightgray;
    background-color: white;
    border-radius: 5px;
    width: 12rem;
    padding: 1rem;
    display: flex;
    flex-direction: column;
    gap: .5rem;
    opacity: 0;
    visibility: hidden;
    transition: opacity 200ms ease-in-out
}


.nav_listitemdrop li {
    list-style: none;
    padding: .5rem 1rem;
    border-radius: 5px;
    transition: background-color 200ms ease-in-out;
}

.nav_listitemdrop li:hover {
    background-color: rgb(13, 218, 207);
}

.spacer {
    height: 4rem;
}
</style>
"@

$pre = @"
<nav class="nav">
<ul class="nav_list">
    <li class="nav_listlogo">
        <a href="https://github.com/gabe-CSX/"><i class='bx bxl-github' ></i></a>
    </li>
    <li class="nav_listitem"><i class='bx bx-user'></i>Users
        <ul class="nav_listitemdrop">
"@
$html = @"
"@
Write-Progress "HTML setup" -Completed
Write-Progress "Run Events"
# If audit disabled, do not run parse events
[bool]$runEvents = $true
$auditLogons = auditpol /get /subcategory:"Logon" /r | ConvertFrom-CSV
if ($auditLogons."Inclusion Setting" -NotMatch 'Success') {
    $html += "<div class='text'><h1>Auditing Logons is not enabled. Consider auditing logins for finding stale accounts.</h1></div>"
    #consider setting an error
    $runEvents = $false
}

$users = Get-ADUser -Filter * -Properties "PasswordLastSet", "Enabled", "PasswordNeverExpires", "Description"
if ($runEvents) {
    $events = Get-WinEvent -FilterHashtable @{
        LogName   = 'Security';
        ID        = 4625;
        StartTime = ((Get-Date).AddDays(-14))
    }
}
Write-Progress "Run Events" -Completed

# Escape here if no logs contain a user's SamAccountName. Otherwise create as many <td> as security events
:Outer foreach ($user in $users) {
    $userSAM = $user.SamAccountName
    Write-Progress "Getting groups of $userSAM"
    $membership = Get-ADPrincipalGroupMembership -Identity $user | Where-Object { $_.Name -match 'admin' }
    if ($membership.name) {
        
        $membership = $membership.name -join ", "
        $lastPass = $user.PasswordLastSet
        $enabled = $user.Enabled
        $pwLastSet = $user.PasswordLastSet
        $pwNoExpire = $user.PasswordNeverExpires

        $pre += @"
        <li><a href="$userSAM">$userSAM</a></li>
"@

        $html += @"
        <div class='user' id="$userSAM">
          <h2>$userSAM</h2>
          <p>Admin memberships: $membership</p>
          <p>Account description: $desc</p>
          <br>
          <table>
            <tr>
              <th>Password last set</th>
              <th>Password never expires</th>
              <th>Account enabled</th>
              <th>Last failed logon time</th>
              <th>Last failed logon type</th>
            </tr>
            <tr>
              <td>$pwLastSet</td>
              <td>$pwNoExpire</td>
              <td>$enabled</td>
"@


        Write-Progress "Getting groups of $userSAM" -Completed
        if ($runEvents) {
            Write-Progress -Activity "Getting events of $userSAM"
            foreach ($log in $events) {
                if ($log.properties[5].value -match "$userSAM") {
                    $logonType = switch ($log.properties[10].value) {
                        2 { 'Interactive' }
                        3 { 'Network' }
                        4 { 'Batch' }
                        5 { 'Service' }
                        7 { 'Unlock' }
                        8 { 'NetworkCleartext' }
                        9 { 'NewCredentials' }
                        10 { 'RemoteInteractive' }
                        11 { 'CachedInteractive' }
                        default {'Unknown'}
                    }
                }
            }
            $log = $log.TimeCreated
            $html += @"
                <td>$log</td>
                <td>$logonType</td>
              </tr>
            </table>
            </div>
"@
    Write-Progress -Activity "Getting events of $userSAM" -Completed
            continue Outer # wow. Getting to here was almost an hour
        }
        
        $html += @"
        <p>No audit log</p>
        <p>NA</p>
        <p>$enabled</p>
        <p>$lastPass</p>

"@ 


    }

}

# Post content
$time = [math]::Round($stopwatch.Elapsed.TotalSeconds, 2)
# successive runs cause issues
#Clear-Variable stopwatch -Scope Global
$pre += @"
</ul>
</li>
<li class="nav_listitem"><a href="#postcontent">Diagnostics</a></li>
</ul>
</nav>
<div class="spacer">
    &nbsp;
</div>
"@
$post = @"
    <hr>
    <div id='postcontent'>
      <h2>Postcontent</h2>
      <p>Powershell execution time in seconds: $time*</p>
      <p>Errors: </p>
        <ul>
"@
if ($Error) {
    foreach ($individual in $Error) {
        $post += "<li>$individual</li>"
    }
} else {
    $post += "<li>No errors in this session.</li>"
}
$post += @"
        </ul>
      <p>*this does not include creating a path, converting to html, and opening the document.</p>
    </div>
"@

# Build and launch HTML
Write-PRogress 'Launching HTML'
$testPath = Test-Path "C:\temp"
if ($testPath -eq $false) {
    Try {
        New-Item -Path "C:\" -Name "temp" -ItemType "Directory" -ErrorAction Stop
    }
    Catch {
        throw "$testPath did not exist, and failed to be created"
    }
}

ConvertTo-Html -Head $head -Body "$pre $html" -PostContent $post | Out-File C:\temp\memberships.html
Start-Process C:\temp\memberships.html

$stopwatch.Stop()
$timeTotal = [math]::Round($stopwatch.Elapsed.TotalSeconds, 2)
Write-Host "Total time $timeTotal"