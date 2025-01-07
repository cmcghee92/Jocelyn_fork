# Project Compassion Powershell Project

## Overview
This project is a PowerShell script designed to help non-profit organizations analyze donations and generate custom notification letters for current or upcoming donation campaigns that can be mailed to previous donors.

## Features
- **Autogenerate custom letters**: Automatically generate personalized thank-you letters for donors.  
- **Integrate with Online Platforms**: Fetch and process donation data from online platforms using APIs.

## Prerequisites
- PowerShell 5.1 or later
- SMTP server credentials for sending emails
- csv file of donor names and emails pulled from a paympent platform.  I used Paypal.

## Installation
1. Clone the repository:
    ```sh
    git clone https://github.com/yourusername/non-profit-donation-tracker.git
    cd non-profit-donation-tracker
    ```

## Usage
1.  For non-profits that want to automatically send a personalized letter that can be mailed to donors.
2.  Customize the script to send reminder emails to previous donors of an upcoming charity event.

### Disclaimer
1.  The PayPal Data.csv file contains movie characters or pop singer names (for fun) and fake addresses.  While I did download the Activity log from Paypal, I only used some of the headers.  None of this data actually came from PayPal to ensure no personal donor information is used.
2.  This is a project to practice use of Powershell scripts and commands in Active Directory to demonstrate knowledge and understanding of powershell commands.
