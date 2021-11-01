
$computers = import-csv -HEADER "computers" -Encoding ASCII -Path C:\viper\AD_SCCM_Recon.csv
$computers | Foreach { 

$actualname = $_.computers 


Get-ADComputer -filter 'Name -like $actualname' -Property info,description   | Select name,enabled,info,description | export-csv -Append -NoTypeInformation -Force -Path C:\Viper\ComputerInformation.csv }

