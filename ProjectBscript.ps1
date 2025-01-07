#ProjectBScript

# Define the path to the CSV file
$csvPath = "C:\Powershell Project\Project Compassion Paypal Data.csv"
# Load the data from the CSV file
$donations = Import-Csv -Path $csvPath

# Parse the DonationDate column as DateTime
$donations | ForEach-Object {
    $_.DonationDate = [datetime]::ParseExact($_.DonationDate, "MM/dd/yyyy", $null)
}

# Sort the data by DonationDate
$donations = $donations | Sort-Object DonationDate

# Group the data by Month and Year, and sum the donations
$groupedDonations = $donations | Group-Object {
    $_.DonationDate.ToString("yyyy-MM")
} | ForEach-Object {
    [PSCustomObject]@{
        MonthYear = $_.Name
        TotalDonation = ($_.Group | Measure-Object DonationAmount -Sum).Sum
    }
}

# Define the path for the output image
$outputImagePath = "C:\Path\To\Your\Donations_Line_Graph.png"

# Generate the line graph
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
foreach ($data in $groupedDonations) {
    $series.Points.AddXY($data.MonthYear, $data.TotalDonation)
}

# Set chart title and axis labels
$chart.Titles.Add("Donations Over Time")
$chartArea.AxisX.Title = "Date"
$chartArea.AxisY.Title = "Donation Amount"

# Save the chart as an image file
$chart.SaveImage($outputImagePath, [System.Windows.Forms.DataVisualization.Charting.ChartImageFormat]::Png)

Write-Host "Line graph has been created successfully at $outputImagePath"
