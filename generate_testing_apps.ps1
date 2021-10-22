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

Function OpenWordDoc {
  param(
    [string]$Filename
  )
  $Word=NEW-Object –comobject Word.Application
  Return $Word.documents.open($Filename)
}

Function SearchAWord { 
  param(
    [object]$Document,
    [string]$findText,
    [string]$replaceWithText
  )
  $FindReplace=$Document.ActiveWindow.Selection.Find
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
  $read_only = $false;
  $visible = $true;
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

ForEach ($student in $studentTable) {
  $fullName = $student.'first name' + ' ' + $student.'last name'
  $belt = $student.'current ranks'
  $class = $student.programs
  $studentNum = $student.pin

  Write-Host "Name: $($fullName)"
  Write-Host "ID: $($studentNum)"
  if ($class -eq 'Little Dragons') {
    Write-Host "Form: $($class)"
    # $Doc=OpenWordDoc -Filename "little_dragon.docx"
    # SearchAWord –Document $Doc -findtext 'something' -replacewithtext 'anotherthing'
    # SaveAsWordDoc –document $Doc –Filename $Savename
  }
  else {
    Write-Host "Form: $($belt)"
    # $Doc=OpenWordDoc -Filename "$belt.docx"
    # SearchAWord –Document $Doc -findtext 'something' -replacewithtext 'anotherthing'
    # SaveAsWordDoc –document $Doc –Filename $Savename
  }
  Write-Host ""
}