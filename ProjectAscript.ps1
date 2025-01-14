#Download weareprojectcompassion.ps1 from GitHub onto C drive then run the script
#Create a powershell script that creates a folder in C drive (in lab) named Powershell Project and downloads the files (paypal data and ps1 script) from GitHub into the folder
#Open powershell ISE as administrator and input script if can't download and run the script in powershell
#open ADUC use View, Advanced Properties, properties, Object and unselect protect object if you need to deleted OU


#Project A-Step 1: Create Donor OU and 5 AD users in Donor OU with email addresses, create the environment

# Create the Donor OU
New-ADOrganizationalUnit -Name "Donor" -Path "DC=Adatum,DC=com"

# Define the base Domain Name for the new users
$DonorOU = "OU=Donor,DC=Adatum,DC=com"

# Define the path to the CSV file
$csvPath = "C:\Powershell Project\Project Compassion Paypal Data.csv"

# Import the CSV file
$userData = Import-Csv -Path $csvPath

# Define the Organizational Unit (OU) where the users will be added
$ouPath = "OU=Donor,DC=Adatum,DC=com"

# Loop through each user in the CSV file
foreach ($user in $userData) {
    # Extract user information from the CSV file
    $firstName = $user.FirstName
    $lastName = $user.LastName
    $userName = $user.FullName
    $email = $user.Email
    $streetAddress = $user.StreetAddress
    $city = $user.City
    $state = $user.State
    $postalCode = $user.PostalCode

# Create the user in Active Directory
    New-ADUser `
        -Name "$firstName $lastName" `
        -GivenName $firstName `
        -Surname $lastName `
        -SamAccountName $userName `
        -UserPrincipalName "$userName@Adatum.com" `
        -Path $ouPath `
        -AccountPassword (ConvertTo-SecureString "Pa55w.d" -AsPlainText -Force) `
        -Enabled $true `
        -EmailAddress $email `
        -StreetAddress $streetAddress `
        -City $city `
        -State $state `
        -PostalCode $postalCode `
        -PassThru | Set-ADUser -ChangePasswordAtLogon $true

    Write-Host "User $firstName $lastName added successfully."
}

Write-Host "All users have been added successfully."


#Project A step-2 pull the list of Donor first and last name and address form the AD OU Donor and create a CSV file

# Define the path to the CSV file
$csvPath = "C:\Powershell Project\Donor Address List.csv"

# Define the Organizational Unit (OU) to search for users
$ouPath = "OU=Donor,DC=Adatum,DC=com"

# Get users from the specified OU with the needed properties
$users = Get-ADUser -Filter * -SearchBase $ouPath -Property GivenName, Surname, StreetAddress, City, State, PostalCode

# Create a list to hold the user data
$userList = @()

# Loop through each user and collect the required information
foreach ($user in $users) {
    $userObject = New-Object PSObject -Property @{
        "FullName"      = "$($user.GivenName) $($user.Surname)"
        "StreetAddress" = $user.StreetAddress
        "City"          = $user.City
        "State"         = $user.State
        "PostalCode"    = $user.PostalCode
    }
    $userList += $userObject
}

# Export the user data to a CSV file with the specified column order 
$userList | Select-Object FullName, StreetAddress, City, State, PostalCode | Export-Csv -Path $csvPath -NoTypeInformation 

Write-Host "CSV file has been created successfully at $csvPath"



#With the Donor Address List.csv autogenerate a custom letter notifying these previous donors of the upcoming charity event thanking them for their donation to Project compassion in the #past and hoping they can donate to the upcoming charity event for a Peruvian Children's hospital treating needy children with traumatic injuries.  Add their Address on the top and add a #custom greeting with their fullName

# Define the path to the CSV file
$csvPath = "C:\Powershell Project\Donor Address List.csv"

# Define the path where the letters will be saved
$outputFolderPath = "C:\Powershell Project\Thank You Letters"

# Create the output folder if it doesn't exist
if (-not (Test-Path -Path $outputFolderPath)) {
    New-Item -ItemType Directory -Path $outputFolderPath
}

# Import the CSV file
$donors = Import-Csv -Path $csvPath

# Loop through each donor in the CSV file
foreach ($donor in $donors) {
    # Extract donor information from the CSV file
    $fullName = $donor.FullName
    $streetAddress = $donor.StreetAddress
    $city = $donor.City
    $state = $donor.State
    $postalCode = $donor.PostalCode
    
    # Define the letter content
    $letterContent = @"
$streetAddress
$city, $state $postalCode

Dear $fullName,

We hope this letter finds you well. We are writing to express our heartfelt gratitude for your generous donation to Project Compassion in the past. 
Your support has made a significant impact on the lives of those in need.

We are excited to announce our current donation drive to raised $5,000 in support of the National Institute of Child Health (INSN) in Lima Peru, 
dedicated to treating needy children with traumatic injuries. Your past contributions have been invaluable, 
and we hope that you will consider donating again to this worthy cause.

The donation drive runs till 3 March 2025 and 100% of all raised funds will go towards paying for medication and treatments for children in need. 
We deeply appreciate your continued support and generosity.

Thank you once again for your kindness and compassion. Together, we can make a difference in the lives of these children.

Warm regards,

Project Compassion
www.weareprojectcompassion.org
"@

    # Define the path for the letter file
    $letterFilePath = "$outputFolderPath\$($fullName.Replace(' ', '_')).txt"

    # Save the letter content to a text file
    $letterContent | Out-File -FilePath $letterFilePath -Encoding UTF8

    Write-Host "Letter for $fullName has been created successfully."
}

Write-Host "All letters have been created successfully."

