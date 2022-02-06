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
:root {
  --background: #fbfbfb;
  --nav: #fbfbfb;
  --highlight: #0aa49b;
  --highlight-text: #fbfbfb;
  --shadow: #bebebe;
  --logoFG: white;
  --logoBG: black;
}

[data-theme="dark"] {
  --background: #212529;
  --nav: #212529;
  --text: #fbfbfb;
  --highlight: #4d67a9;
  --highlight-text: #fbfbfb;
  --shadow: #000;
  --logoFG: black;
  --logoBG: white;
}

* {
  padding: 0;
  margin: 0;
}

html {
  font-family: Arial, Helvetica, sans-serif;
  font-size: 18px;
  width: 100%;
  margin: 0 auto;
  background-color: var(--background);
  color: var(--text);
  scroll-padding-top: 3em;
  scroll-behavior: smooth;
}

.user {
  padding: 1.5rem 0;
  text-indent: 0.75rem;
}

.spacer {
  height: 4rem;
}

h2 {
  color: var(--highlight);
  font-weight: bolder;
}

table {
  margin: 0 auto;
  width: 100%;
  border-spacing: 0;
  text-align: center;
}

th {
  color: var(--highlight-text);
  background-color: var(--highlight);
  padding: 0.125em;
}

table th+th {
  border-left: 1px solid var(--highlight-text);
}

#postcontent {
  padding: 0.75rem;
  list-style-type: circle;
  list-style-position: inside;
}

#postcontent li {
  padding: 1rem;
}

.nav {
  position: fixed;
  width: 100%;
  box-shadow: 0 0 10px var(--shadow);
  color: black;
  background-color: var(--highlight-text);
}

a {
  text-decoration: none;
  color: var(--linkOverride);
}

.nav_list {
  display: flex;
  justify-content: flex-end;
  align-items: center;
  gap: 2rem;
  background-color: var(--nav);
}

.nav_listlogo {
  list-style: none;
  margin-right: auto;
  cursor: pointer;arncho
}

.nav_listlogo i {
  overflow: hidden;
  transition: 200ms ease;
  color: var(--logoFG);
  background-color: var(--logoBG);
  margin: 0 1rem;
  padding: 0.5rem 1rem;
  border-radius: 5px;
  text-align: right;
}

.nav_listlogo i:hover {
  background-color: var(--highlight);
}

.nav_listitem {
  list-style: none;
  font-weight: bold;
  position: relative;
  padding: 1.5rem 1rem;
  cursor: pointer;
  transition: background-color 1s ease-in-out;
  color: var(--text);
}

.nav_listitem:last-child {
  margin: auto 1rem auto auto;
}

.nav_listitem::after {
  content: "";
  width: 80%;
  height: 0.3rem;
  border-radius: 5px;
  position: absolute;
  left: 1rem;
  bottom: 0.8rem;
  background-color: var(--highlight);
  opacity: 0;
  transition: opacity 200ms ease-out;
}

.nav_listitem:hover::after {
  opacity: 1;
}

.nav_listitem:hover .nav_listitemdrop {
  opacity: 1;
  visibility: visible;
}

.nav_listitemdrop {
  position: absolute;
  top: 4rem;
  left: -1rem;
  box-shadow: 0 0 10px var(--shadow);
  background-color: var(--nav);
  border-radius: 5px;
  width: 12rem;
  padding: 1rem;
  display: flex;
  flex-direction: column;
  gap: 0.5rem;
  opacity: 0;
  visibility: hidden;
  transition: opacity 200ms ease-in-out;
}

.nav_listitemdrop li {
  list-style: none;
  padding: 0.5rem 1rem;
  border-radius: 5px;
  transition: background-color 200ms ease-in-out;
  color: var(--text);
}

.nav_listitemdrop li:hover {
  background-color: var(--highlight);
} 

input[type="checkbox"] {
  display: inline-block;
}

label {
  font-weight: bold;
  color: var(--text);
}
</style>
"@

$pre = @"
<nav class="nav">
<ul class="nav_list">
    <li class="nav_listlogo">
        <a href="https://github.com/gabe-CSX/"><i class='bx bxl-github' ></i></a>
    </li>
    <div>
      <label>
        <input type="checkbox" id="checkbox"/> Dark mode
      </label>
      <script>
        const toggleSwitch = document.getElementById('checkbox')

        function switchTheme() {
            if (toggleSwitch.checked) {
                document.documentElement.setAttribute('data-theme', 'dark');
                localStorage.setItem('theme', 'dark');
            }
            else {
                document.documentElement.setAttribute('data-theme', 'light');
                  localStorage.setItem('theme', 'light');
              }    
        }
      
        toggleSwitch.addEventListener('change', switchTheme, false);
      </script>
    </div>
    <li class="nav_listitem"><i class='bx bx-user'></i>Users
        <ul class="nav_listitemdrop">
"@
$html = @"
"@
Write-Progress "HTML setup" -Completed

Write-Progress "Run Events"
$users = Get-ADUser -Filter * -Properties "PasswordLastSet", "Enabled", "PasswordNeverExpires", "Description"
$events = Get-WinEvent -FilterHashtable @{
  LogName   = 'Security';
  ID        = 4625;
  StartTime = ((Get-Date).AddDays(-14))
}
Write-Progress "Run Events" -Completed

foreach ($user in $users) {
  $userSAM = $user.SamAccountName
  Write-Progress "Getting groups of $userSAM"
  $membership = Get-ADPrincipalGroupMembership -Identity $user | Where-Object { $_.Name -match 'admin' }
  if ($membership.name) {
        
    $membership = $membership.name -join ", "
    $enabled = $user.Enabled
    $pwLastSet = $user.PasswordLastSet
    $pwNoExpire = $user.PasswordNeverExpires

    $pre += @"
    <a class='anchor' href="#$userSAM"><li>$userSAM</li></a>
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

    Write-Progress -Activity "Getting events of $userSAM"
    if ($events) {
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
            default { 'Unknown' }
          }
        }
      }
      $log = $log.TimeCreated
    }
    else {
      $log = "No logs found. Audit events may be disabled."
      $logonType = "NA"
    }
    
    $html += @"
                <td>$log</td>
                <td>$logonType</td>
              </tr>
            </table>
            </div>
"@
    Write-Progress -Activity "Getting events of $userSAM" -Completed



  }

}


# Post content
$time = [math]::Round($stopwatch.Elapsed.TotalSeconds, 2)
$pre += @"
        </ul>
      </li>
      <a class='anchor' href="#postcontent"><li class="nav_listitem">Diagnostics</li></a>
  </ul>
</nav>
<div class="spacer">
  &nbsp;
</div>
"@

# This takes a while and I'd like to continue recording the time
$path = 'C:\temp\DefaultAssoc.xml'
Dism.exe /online /Export-DefaultAppAssociations:$path
$htmlApp = (Select-Xml -Path $path -xpath '/DefaultAssociations/*' | ForEach-Object { $_.node } | Where-Object { $_.Identifier -eq '.html' }).ApplicationName
if ($htmlApp -match 'Internet Explorer') {
  Write-Error 'Opened in Internet Explorer. Please open in any modern browser.'
}


# Timestamp
$post = @"
    <hr>
    <div id='postcontent'>
      <h2>Postcontent</h2>
      <p>Powershell execution time in seconds: $time*</p>
      <p>Errors: </p>
        <ul>
"@


# Append errors and suggestions
if ($Error) {
  foreach ($individual in $Error) {
    $post += "<li>$individual</li>"
  }
}
else {
  $post += "<li>No errors in this session.</li>"
}
$post += @"
        </ul>
      <p>Common errors:<p>
        <ul>
          <li>Cannot find security log/Parameter incorrect: Run interactively. Check membership to event log readers and registry key.</li>
          <li>
            No events were found matching the criteria: Use 'auditpol /get /subcategory:logon'. If failures are not logged, consider enabled per best practice.
            <a style='text-decoration: unline; color: blue' href='https://docs.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/audit-policy-recommendations'> Security Best Practices</a>.
          </li>
      <p>*this does not include creating a path, converting to html, and opening the document.</p>
    </div>
"@


# Build and launch HTML
Write-Progress 'Launching HTML'
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
Start-Process C:/temp/memberships.html

$stopwatch.Stop()
$timeTotal = [math]::Round($stopwatch.Elapsed.TotalSeconds, 2)
Write-Host "Total time $timeTotal"
