# Obesogenic Diet Brain Study - Lab Results Tracker

Web application for recording and visualizing lab results from drug studies investigating physiological brain changes caused by obesogenic diets in rats.

## Features

- Multi-user authentication (login/register)
- Data entry for rat subject identifiers, drug/compound info, and brain regions
- Searchable, sortable results table
- 6 interactive charts (weight comparisons, brain regions, drug distribution, etc.)
- Excel export

## Setup

### Prerequisites

- [Node.js](https://nodejs.org/) (v18 or later)

### Installation

1. Download or clone this folder
2. Open a terminal and navigate to the project folder:
   ```bash
   cd lab-results-app
   ```
3. Install dependencies:
   ```bash
   npm install
   ```
4. Start the server:
   ```bash
   node server.js
   ```
5. Open **http://localhost:3000** in your browser

### First Login

A default admin account is created on first run:
- **Username:** admin
- **Password:** admin123

Anyone can create their own account by clicking "Create one" on the login screen.

## Data

All data is stored locally in a SQLite database (`lab_data.db`) that is created automatically on first run. Each user's machine will have its own independent database.
