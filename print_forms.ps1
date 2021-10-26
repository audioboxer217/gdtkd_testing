$workingDir = Get-Location
$outputDir = "${workingDir}/testing_forms"

$formList = Get-ChildItem -Path $outputDir

foreach ($form in $formList) {
    Write-Host $form
}