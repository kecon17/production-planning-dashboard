# Weekly Production Planning Dashboard

## Overview

This is a web-based application designed for production planners and shop-floor managers to optimize weekly operational schedules. The tool reads demand, station, time, training, and product data from Excel files, runs a scheduling algorithm to generate multiple weekly production proposals, and visualizes the results in an interactive dashboard.

The primary goal is to provide a clear, data-driven plan that respects operator training, work center constraints, and daily working hours, helping to maximize throughput and resource utilization.

## Features

-   **Weekly-Focused Planning:** All views and logic are centered around a full Monday-to-Friday production week.
-   **Dynamic Operator Availability:** Planners can configure which operators are available for each specific day of the week, allowing for flexible scheduling around absences.
-   **Automated ETL:** A single click generates three distinct weekly production proposals based on current data and a heuristic-based scheduling algorithm.
-   **Interactive Gantt Chart:** A large-format Gantt chart visualizes the entire weekly schedule, showing which operator is assigned to which task, in which work center, and at what time.
-   **Production Summary:** A clear summary table shows the total count of each subsystem scheduled for production within the proposed week.
-   **Demand Analysis:** A separate tab provides tools to analyze the monthly and weekly demand targets, helping to contextualize the operational plan.

## Project Structure

```
planning/
│
├── data/
│   ├── demand.xlsx
│   ├── products.xlsx
│   ├── stations.xlsx
│   ├── times.xlsx
│   └── training.xlsx
│
├── output/
│   └── planning_week_YYYYMMDD_proposalN.csv  (Generated files)
│
├── etl.py                  # The backend scheduling and data processing logic.
├── dashboard.py            # The Streamlit frontend application.
├── README.md               # This file.
└── requirements.txt        # Python dependencies.
```

## Setup and Installation

Follow these steps to set up and run the project locally.

#### Prerequisites

-   Python 3.9 or higher
-   Git installed on your system

#### 1. Clone the Repository

Open your terminal or command prompt and clone the project to your local machine (you will do this after creating the GitHub repository).
```bash
git clone <your-github-repository-url>
cd planning
```

#### 2. Create a Virtual Environment

It is highly recommended to use a virtual environment to manage project dependencies.
```bash
# For Windows
python -m venv venv
.\venv\Scripts\activate

# For macOS/Linux
python3 -m venv venv
source venv/bin/activate
```

#### 3. Install Dependencies

Install all the required Python libraries using the `requirements.txt` file.
```bash
pip install -r requirements.txt
```

#### 4. Configure Data Files

-   Create a folder named `data` inside the main `planning` directory.
-   Place the following required Excel files into the `planning/data/` folder:
    -   `demand.xlsx`
    -   `products.xlsx`
    -   `stations.xlsx`
    -   `times.xlsx`
    -   `training.xlsx`
-   Ensure the `DATA_DIR` path variable at the top of `etl.py` points to the correct location of your data files if you choose to place them elsewhere.

## How to Use

#### 1. Run the Dashboard

Open your terminal, make sure your virtual environment is activated, and run the following command from the `planning` directory:
```bash
streamlit run dashboard.py # option 1
python -m streamlit run dashboard.py # option 2
```
Your web browser should automatically open with the application running.

#### 2. Application Workflow

1.  **Configure Availability:** In the sidebar, expand the sections for each day of the week and select the operators who will be available.
2.  **Generate Proposals:** Click the "Generate Weekly Proposals" button. The ETL process will run in the background. Once complete, the app will refresh.
3.  **Select a Proposal:** Choose one of the generated proposals (e.g., "Proposal 1") from the radio buttons at the top of the "Planning" tab.
4.  **Analyze the Plan:** Review the "Weekly Plan" table, the interactive "Weekly Gantt Chart", and the "Weekly Production Summary" to understand the proposed schedule.
5.  **Analyze Demand:** Switch to the "Demand" tab to view the underlying monthly and weekly demand data that drives the plan.

## Future Enhancements

This tool provides a strong foundation for operational planning. Future enhancements could include:

1.  **Live Data Integration:** Connect to ERP/MES and HR systems to automate the import of demand, bill of materials, and operator availability, eliminating manual data handling.
2.  **Constraint Optimization:** Replace the current heuristic scheduler with a true constraint solver (like Google OR-Tools) to find mathematically optimal plans based on business goals (e.g., minimizing lateness, maximizing throughput).
3.  **"What-If" Scenarios:** Add a simulation mode to allow planners to model the impact of rush orders, machine downtime, or changes in operator availability before committing to a plan.