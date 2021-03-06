Param(
    [string]$studentFile,
    [string]$testDate,
    [string]$dueDate,
    [string]$schoolName='goldendragontkdowasso',
    [string]$username=$env:GDTKD_USERNAME,
    [string]$password=$env:GDTKD_PASSWORD
)

Function getAuthDetails {
  $loginUri = "https://id.kicksite.net/$schoolName"
  $loginPage = Invoke-WebRequest  -Uri $loginUri -SessionVariable session
  $fields = $loginPage.InputFields
  ForEach ($field in $fields | Where-Object {$_.Name -eq 'authenticity_token'}) {
    $token = $field.value
  }

  Return @{token=$token; session=$session}
}

Function login {
  param (
    [string]$token,
    [object]$session
  )
  $userSessionUri = 'https://id.kicksite.net/user_sessions'
  $form = @{
    utf8='✓'
    authenticity_token=$token
    custom_subdomain=$schoolName
    username=$username
    password=$password
    kiosk='false'
    commit='LOG IN'
  }

  Invoke-WebRequest -Method 'Post' -WebSession $session -Form $form -Uri $userSessionUri
}

Function GetStudentFile {
  param (
    [object]$session
  )
  $kicksiteURL = "https://goldendragontkdowasso.kicksite.net/students.csv?sort_by=created_at&sort_direction=desc&status=active_without_frozen"

  Invoke-WebRequest -WebSession $session -Outfile studentFile.csv -Uri $kicksiteURL
}

Function GetNextBelt {
  param(
    [string]$belt,
    [string]$class
  )

    $beltCollection = @("White Belt")

    if ($class -eq 'Little Dragons') {
      $beltCollection += @(
        "White w/ Yellow Stripe",
        "White w/ Orange Stripe",
        "White w/ Green Stripe",
        "White w/ Blue Stripe",
        "White w/ Red Stripe"
      )
    }

    $beltCollection += @(
      "Yellow Stripe",
      "Yellow Belt",
      "Orange Stripe",
      "Orange Belt",
      "Green Stripe",
      "Green Belt",
      "Blue Stripe",
      "Blue Belt",
      "Red Stripe",
      "Red Belt",
      "Brown Stripe",
      "Brown Belt",
      "Black Stripe"
    )

    $currBeltIndex = $beltCollection.IndexOf($belt)

    return $beltCollection[$currBeltIndex + 1]

}

Function OpenWordDoc {
  param(
    [string]$Filename
  )
  $Word = New-Object -ComObject Word.Application
  Return $Word.documents.open($Filename)
}

Function SearchAWord { 
  param(
    [object]$Document,
    [string]$findText,
    [string]$replaceWithText
  )
  $FindReplace = $Document.ActiveWindow.Selection.Find
  $matchCase = $false;
  $matchWholeWord = $true;
  $matchWildCards = $false;
  $matchSoundsLike = $false;
  $matchAllWordForms = $false;
  $forward = $true;
  $format = $false;
  $matchKashida = $false;
  $matchDiacritics = $false;
  $matchAlefHamza = $false;
  $matchControl = $false;
  # $read_only = $false;
  # $visible = $true;
  $replace = 2;
  $wrap = 1;
  $FindReplace.Execute($findText, $matchCase, $matchWholeWord, $matchWildCards, $matchSoundsLike, $matchAllWordForms, $forward, $wrap, $format, $replaceWithText, $replace, $matchKashida ,$matchDiacritics, $matchAlefHamza, $matchControl)
}

Function SaveAsWordDoc {
  param(
    [object]$Document,
    [string]$Filename
  )
  $Document.Saveas([REF]$Filename)
  $Document.close()
}

if (!$studentFile) {
  $authDetails = getAuthDetails
  $session = login @authDetails
  GetStudentFile $authDetails.session
  $studentFile = 'studentFile.csv'
}
if (!$testDate) {
  $testDate = Read-Host -Prompt "Testing Date"
}
if (!$dueDate) {
  $dueDate = Read-Host -Prompt "Due Date"
}

$studentTable = Import-Csv -Path $studentFile
$studentList = $studentTable | Where-Object {$_.'current ranks' -notlike '*Black*'}

$beltOrders = @{}

$workingDir = Get-Location
$outputDir = "${workingDir}/testing_forms"
New-Item -ItemType Directory -Force -Path "$outputDir" | Out-Null

ForEach ($student in $studentList) {
  $fullName = $student.'first name' + ' ' + $student.'last name'
  $class = $student.programs
  $belt = $student.'current ranks'
  $nextBelt = GetNextBelt $belt $class
  $beltSize = $student.'belt size'
  $studentNum = $student.pin
  $studentAge = $student.age

  if (!$beltOrders[$nextBelt]) {
    $beltOrders[$nextBelt] = @{}
  }

  if (!$beltOrders[$nextBelt][$beltSize]) {
    $beltOrders[$nextBelt][$beltSize] = 1
  } 
  else {
    $beltOrders[$nextBelt][$beltSize] += 1
  }

  Write-Host "Name: $($fullName)"
  Write-Host "ID: $($studentNum)"

  if ($class -eq 'Little Dragons') {
    Write-Host "Form: $($class)"
    Write-Host "Next Belt: $($nextBelt)"
    $nextBeltClean = $nextBelt.replace('w/','+')

    $Doc = OpenWordDoc -Filename "${PSScriptRoot}\templates\Little Dragons.docx"
    SearchAWord -Document $Doc -findtext 'TEST_DATE' -replacewithtext $testDate | Out-Null
    SearchAWord -Document $Doc -findtext 'DUE_DATE' -replacewithtext $dueDate | Out-Null
    SearchAWord -Document $Doc -findtext 'STUDENT_NAME' -replacewithtext $fullName | Out-Null
    SearchAWord -Document $Doc -findtext 'ID' -replacewithtext $studentNum | Out-Null
    SearchAWord -Document $Doc -findtext 'CURR_BELT' -replacewithtext $belt | Out-Null
    SearchAWord -Document $Doc -findtext 'NEXT_BELT' -replacewithtext $nextBelt | Out-Null
    SearchAWord -Document $Doc -findtext 'BELT_SIZE' -replacewithtext $beltSize | Out-Null
    SaveAsWordDoc -document $Doc -Filename "$outputDir\${fullName}- ${nextBeltClean}.docx"
  }
  else {
    Write-Host "Form: $($nextBelt)"
    Write-Host "Next Belt: $($nextBelt)"

    $Doc = OpenWordDoc -Filename "${PSScriptRoot}\templates\$nextBelt.docx"
    SearchAWord -Document $Doc -findtext 'TEST_DATE' -replacewithtext $testDate | Out-Null
    SearchAWord -Document $Doc -findtext 'DUE_DATE' -replacewithtext $dueDate | Out-Null
    SearchAWord -Document $Doc -findtext 'STUDENT_NAME' -replacewithtext $fullName | Out-Null
    SearchAWord -Document $Doc -findtext 'ID' -replacewithtext $studentNum | Out-Null
    SearchAWord -Document $Doc -findtext 'STUDENT_AGE' -replacewithtext $studentAge | Out-Null
    SaveAsWordDoc -document $Doc -Filename "$outputDir\${fullName} - ${nextBelt}.docx"
  }
  Write-Host ""
}

$orderTotal = 0
foreach ($item in $beltOrders.Keys) {
 Write-Host $item
 $sizes = $beltOrders[$item]
 $orderTotal += ($sizes.Values | Measure-Object -Sum).Sum
 foreach ($size in $sizes.Keys) {
   Write-Host "$($size): $($beltOrders[$item][$size])"
 }
}
Write-Host $orderTotal
