
param(

[string] $UNC


)


if (-not $UNC ){
"NO UNC GIVEN!"
"Usage " + $0 + " -UNC //server/share"
exit}

# Set variables to indicate value and key to set

$Apps=@("Word","Excel","Powerpoint")
$Values=@()

foreach ($App in $Apps){



$RegistryPath = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\' + $App + '\Security\Trusted Locations'



# Now set the value
New-ItemProperty -Path $RegistryPath -Name "allownetworklocations" -Value 1 -PropertyType "DWORD" -Force 



# Create the key if it does not exist
If (-NOT (Test-Path $RegistryPath)) {
  New-Item -Path $RegistryPath -Force | Out-Null
}  

$datum = Get-Date -UFormat "%m/%d/%Y %R"








$RegistryPath+="\Location20"
# Create the key if it does not exist
If (-NOT (Test-Path $RegistryPath)) {
  New-Item -Path $RegistryPath -Force | Out-Null
}  


$Values=@(
("Path", $UNC,"String"),
("AllowSubfolders", "1","DWORD"),

("Description", "","String"),
("Date", $datum,"String")


)


foreach ($Value in $Values){


# Now set the value
New-ItemProperty -Path $RegistryPath -Name $Value[0] -Value $Value[1] -PropertyType $Value[2] -Force 

}

}
