# Hostel Meal Management System

A Python-based application for automating meal tracking, budgeting, and charge calculations for hostels. This system simplifies the management of boarder meals, guest meals, deposits, and expenses using Excel spreadsheets as the data backend.

## What This Project Does

This project provides a comprehensive solution for hostel managers to:

- **Track Daily Meals**: Record and manage boarder meals for both day and night, including special meals like guest meals and grand guest meals.
- **Automate Meal Counting**: Automatically turn on meals for boarders based on predefined rules, time schedules, and individual preferences.
- **Calculate Budgets**: Compute meal charges, establishment charges, and final dues/refunds based on deposits and expenses.
- **Manage Boarders**: Add, edit, and monitor individual boarder information including dietary restrictions, deposit amounts, and accessories charges.
- **Generate Reports**: Produce final summary sheets with detailed charge breakdowns for each boarder.
- **Handle Expenses**: Track marketing costs, extra rice purchases, and establishment expenses.

## Key Features

### Auto Meal Calculation
- Automatically counts and assigns meals based on current time and boarder status.
- Supports rules like dietary restrictions (no beef, no fish, etc.), weekend preferences, and meal timing (day/night only).
- Handles guest meals with configurable pricing.

### Auto Meal Counting
- Tracks total meals per boarder, including regular meals, guest meals, and extra eggs.
- Calculates totals for the entire hostel, including marketing and establishment expenses.
- Provides real-time summaries of meal counts and budget status.

### Budget Management
- Calculates meal charges per boarder based on total expenses divided by total meals.
- Computes establishment charges (shared costs like electricity, maintenance) distributed among boarders.
- Tracks deposits, dues, and refunds with detailed breakdowns.
- Monitors low-budget boarders and provides warnings.

### Excel Integration
- Uses Excel files as the primary data storage and reporting format.
- Generates template sheets for meal counts, marketing, establishment costs, and meal routines.
- Supports color-coded cells for easy visual identification of meal status.

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/s-k-zaman/auto-mealer-hostel
   cd auto-mealer-hostel
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Run the application:
   ```bash
   python main.py
   ```

## Usage

### Initial Setup
1. Run `python main.py init` to create a template Excel file (`hostel-meals.xlsx`) with empty sheets.
2. The system will prompt for month, year, and optionally load boarder names from `boarders.txt`.

### Daily Operations
1. Launch the main menu: `python main.py`
2. Select a boarder to manage individual settings or view details.
3. Use "Do meal count (automatic)" to automatically assign meals for the current day/time.
4. Monitor totals and budgets through the main menu information display.

### Key Menu Options
- **Select a Boarder**: Manage individual boarder settings (deposits, rules, meal status).
- **Set/Remove Grand Meal Day**: Mark special days with higher guest meal pricing.
- **Add a Boarder**: Add new residents to the system.
- **See Meal Count Details**: View meals for specific dates/times.
- **See Deposit/Meal Details**: Check boarder budget status and warnings.
- **Produce Final Sheet**: Generate a comprehensive charge summary sheet.

### Boarder Management
For each boarder, you can:
- Update deposits and accessories charges.
- Toggle meal status (on/off/continue).
- Set dietary rules (no beef, weekends off, etc.).
- Manually turn meals on/off for specific dates.
- Add extra eggs or guest meals.

## How It Helps

### Auto Calculation
- **Meal Charges**: Automatically calculates per-meal cost by dividing total expenses (marketing + extra rice + extra marketing) minus guest meal charges by total meals consumed.
- **Establishment Charges**: Distributes shared costs (electricity, maintenance) equally among boarders.
- **Dues and Refunds**: Computes final amounts owed or refundable based on deposits vs. calculated charges.

### Auto Counting Meals
- **Real-time Tracking**: Counts meals as they are marked, providing instant totals.
- **Guest Meal Handling**: Automatically includes guest meals in total counts with appropriate pricing.
- **Rule-based Automation**: Applies individual boarder rules to automatically set meal status without manual intervention.

### Budget Management
- **Expense Tracking**: Monitors all hostel expenses including marketing, extra purchases, and establishment costs.
- **Budget Alerts**: Warns when boarders have low deposits relative to expected charges.
- **Final Reconciliation**: Produces detailed statements showing exactly how charges are calculated and what amounts are due or refundable.

## Configuration

Key settings in `config.py`:
- Meal pricing (regular, guest, grand guest, eggs)
- Time thresholds for meal counting
- Default charges and thresholds
- Manager name and currency symbol
