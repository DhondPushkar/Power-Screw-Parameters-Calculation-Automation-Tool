# Power-Screw-Parameters-Calculation-Automation-Tool
Power Screw Parameter Calculation Automation Tool using Python, Excel

## Project Overview

This project is a Python-based academic automation tool developed to assist faculty members and researchers working on mechanical engineering research related to **power screw performance analysis**.

During research work, faculty were required to repeatedly calculate multiple power screw parameters (torque, efficiency, stresses, etc.) for different input values and manually enter the results into Excel sheets. This process was **time-consuming, repetitive, and prone to manual errors**.

This tool eliminates that manual effort by embedding all standard power screw formulas into a reusable Python application. Faculty can simply provide input parameters (manually or via Excel), and the tool automatically computes all required results and saves them back to an Excel file, ready to be used directly in research papers.

---

## Problem Statement

* Manual calculation of power screw parameters for multiple data points
* High chances of calculation and transcription errors
* Significant time spent entering computed values into Excel
* Reduced research productivity due to repetitive work

---

## Solution

A Python-based automation tool that:

* Stores all required power screw formulas
* Accepts input parameters through an Excel file
* Performs bulk calculations automatically
* Exports all calculated parameters into a structured Excel output

This removes the need for manual calculations and manual Excel entry.

---

## Key Features

* Bulk calculation using Excel input
* Automated computation of all major power screw parameters
* Clean and simple Tkinter-based user interface
* One-click Excel export for research documentation
* Accurate and consistent results across multiple datasets

---

## Input Parameters (Excel Columns Required)

The input Excel file must contain the following columns:

* `Pitch`
* `Threads`
* `NominalDiameter`
* `CoreDiameter`
* `FrictionAngle` (in degrees)
* `Load`

Each row represents one power screw data point.

---

## Calculated Output Parameters

For each input row, the tool calculates:

* Mean Diameter
* Lead
* Helix Angle
* Torque required to raise the load
* Torque required to lower the load
* Efficiency
* Maximum Efficiency
* Overall Efficiency
* Shear Stress
* Compressive Stress
* Maximum Principal Stress
* Maximum Shearing Stress

All results are automatically saved into a new Excel file.

---

## Technology Stack

* **Python** – Core logic and automation
* **Tkinter** – Graphical user interface
* **Pandas** – Excel data handling and bulk processing
* **Math Library** – Mathematical and trigonometric calculations

---

## How It Works

1. Faculty prepares an Excel file containing input parameters
2. User launches the application
3. Excel file is uploaded through the UI
4. Tool automatically performs all calculations
5. Results are saved into a new Excel file
6. Output file can be directly used in research papers

---

## Use Case

* Mechanical engineering academic research
* Power screw performance analysis
* Research paper data preparation
* Engineering calculation automation

---

## Learning Outcomes

Through this project, I gained hands-on experience in:

* Converting domain-specific mathematical formulas into executable logic
* Automating repetitive engineering calculations
* Working with structured Excel data using Python
* Building user-friendly academic tools
* Problem-solving and requirement understanding

---

## Project Positioning

This is an **academic automation project**, not a data analytics or machine learning project. It demonstrates:

* Problem-solving ability
* Python programming fundamentals
* Mathematical modeling
* Automation of real-world workflows

---

## Author

Developed by **Pushkar Dhond** during academic coursework to support faculty research work.

---

## Future Enhancements

* Manual parameter entry alongside Excel upload
* Input validation and error handling
* Result visualization (charts)
* Export to CSV and PDF formats

---

## Disclaimer

This tool was developed for academic and educational purposes to assist research activities.

