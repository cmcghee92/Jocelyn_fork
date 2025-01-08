#ProjectBScript

# Define the path to the CSV file
$csvPath = "C:\Powershell Project\Project Compassion Paypal Data.csv"
# Load the data from the CSV file
$donations = Import-Csv -Path $csvPath

# Parse the DonationDate column as DateTime with error handling
$donations | ForEach-Object {
    try {
        $_.DonationDate = [datetime]::ParseExact($_.DonationDate, "MM/dd/yyyy", $null)
    } catch {
        Write-Warning "Failed to parse date: $($_.DonationDate)"
    }
}

# Sort the data by DonationDate
$donations = $donations | Sort-Object DonationDate

# Filter donations by year
$donations2021 = $donations | Where-Object { $_.DonationDate.Year -eq 2021 }
$donations2022 = $donations | Where-Object { $_.DonationDate.Year -eq 2022 }
$donations2023 = $donations | Where-Object { $_.DonationDate.Year -eq 2023 }

# Function to generate line graph
function Generate-LineGraph {
    param (
        [array]$data,
        [string]$outputPath,
        [string]$title
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Windows.Forms.DataVisualization

    $chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
    $chart.Width = 800
    $chart.Height = 600

    $chartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $chart.ChartAreas.Add($chartArea)

    $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series
    $series.Name = "Donations"
    $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
    $chart.Series.Add($series)

    # Add data points to the series
    $months = @("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    foreach ($month in $months) {
        $dataPoint = $data | Where-Object { $_.MonthYear -like "*$month*" }
        if ($dataPoint) {
            $series.Points.AddXY($month, $dataPoint.TotalDonation)
        } else {
            $series.Points.AddXY($month, 0)
        }
    }

    # Set chart title and axis labels
    $chart.Titles.Add($title)
    $chartArea.AxisX.Title = "Month"
    $chartArea.AxisY.Title = "Donation Amount"

    # Save the chart as an image file
    $chart.SaveImage($outputPath, [System.Windows.Forms.DataVisualization.Charting.ChartImageFormat]::Png)
}

# Function to generate bar chart
function Generate-BarChart {
    param (
        [array]$data,
        [string]$outputPath,
        [string]$title
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Windows.Forms.DataVisualization

    $chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
    $chart.Width = 800
    $chart.Height = 600

    $chartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $chart.ChartAreas.Add($chartArea)

    $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series
    $series.Name = "Total Donations"
    $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Bar
    $chart.Series.Add($series)

    # Add data points to the series
    foreach ($dataPoint in $data) {
        $series.Points.AddXY($dataPoint.Year, $dataPoint.TotalDonation)
    }

    # Set chart title and axis labels
    $chart.Titles.Add($title)
    $chartArea.AxisX.Title = "Year"
    $chartArea.AxisY.Title = "Total Donation Amount"

    # Save the chart as an image file
    $chart.SaveImage($outputPath, [System.Windows.Forms.DataVisualization.Charting.ChartImageFormat]::Png)
}

# Group and generate line graph for each year
$years = @(2021, 2022, 2023)
foreach ($year in $years) {
    $filteredDonations = $donations | Where-Object { $_.DonationDate.Year -eq $year }
    $groupedDonations = $filteredDonations | Group-Object {
        $_.DonationDate.ToString("yyyy-MM")
    } | ForEach-Object {
        [PSCustomObject]@{
            MonthYear = $_.Name
            TotalDonation = ($_.Group | Measure-Object DonationAmount -Sum).Sum
        }
    }

    # Ensure each month from January to December is listed
    $months = 1..12
    $completeData = @()
    foreach ($month in $months) {
        $monthYear = "{0:D4}-{1:D2}" -f $year, $month
        $donation = $groupedDonations | Where-Object { $_.MonthYear -eq $monthYear }
        if ($donation) {
            $completeData += $donation
        } else {
            $completeData += [PSCustomObject]@{
                MonthYear = $monthYear
                TotalDonation = 0
            }
        }
    }

    $outputImagePath = "C:\Powershell Project\Donations_Line_Graph_$year.png"
    Generate-LineGraph -data $completeData -outputPath $outputImagePath -title "Donations for $year"
}

# Calculate total donations by year
$totalDonationsByYear = $years | ForEach-Object {
    $year = $_
    $totalDonation = ($donations | Where-Object { $_.DonationDate.Year -eq $year } | Measure-Object DonationAmount -Sum).Sum
    [PSCustomObject]@{
        Year = $year
        TotalDonation = $totalDonation
    }
}

# Generate bar chart for total donations by year
$outputBarChartPath = "C:\Powershell Project\Total_Donations_By_Year.png"
Generate-BarChart -data $totalDonationsByYear -outputPath $outputBarChartPath -title "Total Donation by Year"

Write-Host "Line graphs have been created successfully for 2021, 2022, and 2023"
Write-Host "Bar chart has been created successfully for total donations by year"
