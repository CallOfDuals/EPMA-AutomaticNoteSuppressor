# EPMA Drug Suppression Script

This Python script automates the process of managing patient records within the Electronic Prescribing and Medicines Administration (EPMA) system. It was developed during my time working with the EPMA team at the Norfolk and Norwich Hospital, where I was responsible for helping to streamline electronic drug order management.

## Purpose

The script automatically navigates through the EPMA system, searches for patient records by hospital number, and suppresses outdated or unnecessary drug order notes based on predefined criteria. This process involves interacting with the EPMA web interface and managing patient notes efficiently through the use of Selenium for browser automation.

## How It Works

- **Data Input**: The script reads hospital numbers and associated drug orders from a pre-defined Excel file.
- **Login & Navigation**: It logs into the EPMA system using user-provided credentials and navigates to the relevant patient records.
- **Patient Search & Note Suppression**: The script searches for specific patients, checks for existing drug orders, and suppresses any matching notes.
- **Automation**: All actions, from logging in to note suppression, are handled programmatically, minimising the need for manual intervention.

This solution reduced manual data entry and ensured the accuracy of drug suppression notes, saving tens of hours per week for the Pharmacists.
