Do Not use, use ProjectAscript.ps1 or ProjectBscript.ps1


#Project A-Step 1: Create Donor OU and 5 AD users in Donor OU with email addresses

# Create the Donor OU
New-ADOrganizationalUnit -Name "Donor" -Path "DC=Adatum,DC=com"

# Define the base DN for the new users
$DonorOU = "OU=Donor,DC=Adatum,DC=com"

# Create 5 Donor users and add email addresses
for ($i = 1; $i -le 5; $i++) {
    $username = "DonorUser$i"
    $password = "Pa55w.rd"  # Ensure you follow your organization's password policy
    $email = "$username@Adatum.com"
    
    # Create the user
    New-ADUser -Name $username -SamAccountName $username -UserPrincipalName $email `
               -Path $DonorOU -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) `
               -Enabled $true
    
    # Set the email address
    Set-ADUser -Identity $username -EmailAddress $email
}

Write-Host "Donor OU and 5 Donor users with email addresses created successfully."

#Project A-Step 2: Create the CSV file with donor name, donor email, donation amount and donation date

# Define the OU path where the donor users are located
$DonorOU = "OU=Donor,DC=Adatum,DC=com"

# Get the 5 donor users from the specified OU
$donorUsers = Get-ADUser -Filter * -SearchBase $DonorOU -Property EmailAddress | Select-Object *

# Define the donation details (Create an array table with Donor info)
$donationDetails = @(
    [PSCustomObject]@{DonorName = "DonorUser1"; Email = "DonorUser1@Adatum.com"; DonationAmount = 100; DonationDate = "2025-01-01"}
    [PSCustomObject]@{DonorName = "DonorUser2"; Email = "DonorUser2@Adatum.com"; DonationAmount = 150; DonationDate = "2025-01-02"}
    [PSCustomObject]@{DonorName = "DonorUser3"; Email = "DonorUser3@Adatum.com"; DonationAmount = 200; DonationDate = "2025-01-03"}
    [PSCustomObject]@{DonorName = "DonorUser4"; Email = "DonorUser4@Adatum.com"; DonationAmount = 250; DonationDate = "2025-01-04"}
    [PSCustomObject]@{DonorName = "DonorUser5"; Email = "DonorUser5@Adatum.com"; DonationAmount = 300; DonationDate = "2025-01-05"}
)

# Create a new CSV file and export the donation details
$csvFilePath = "C:\donor_donations.csv"
$donationDetails | Export-Csv -Path $csvFilePath -NoTypeInformation

Write-Host "Donation details have been exported to $csvFilePath successfully."

#Project A Step 3- CReate a thank you letter (.txt file) for each Donor in the array table (csv file donor_donations.csv)

# Define the array of donors
$donors = $donationDetails

# Define the output directory for the thank you letters
$outputDir = "C:\DonorThankYouLetters"
if (-not (Test-Path -Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory
}

# Loop through each donor and create a thank you letter
foreach ($donor in $donors) {
    $letterContent = @"
Dear $($donor.DonorName),

Mahalo nui loa for your generous donation of $($donor.AmountDonated) on $($donor.DonationDate). Your support is greatly appreciated and helps us continue our mission to deliver food, clothing, toys and schools supplies to hardworking families in need located in rural towns throughout Peru. 

Since our inception in 2019 donations from incredibly caring people have helped over 200 families and are grateful you chose to be a part of our mission!

Forever grateful,
Project Compassion
https://weareprojectcompassion.org/

"@

    # Define the path for the thank you letter file
    $letterFilePath = "$outputDir\$($donor.DonorName)_ThankYouLetter.txt"

    # Create the thank you letter file
    $letterContent | Out-File -FilePath $letterFilePath
}

Write-Host "Thank you letters have been created successfully in $outputDir."

