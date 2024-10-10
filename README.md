# Class List PDF Generator

This project was created to assist my class rep in automating the process of generating class lists for various subjects. Traditionally, creating lists for attendance, assessments, or assignments can be time-consuming, especially with the frequent requests from lecturers. This script significantly reduces the workload by automating the creation of neatly formatted class lists in PDF format.

## Purpose

The main goal of this project is to streamline the process of generating subject-based class lists, saving valuable time and reducing the manual effort involved. By automating this task, the class rep can quickly provide the required lists, making it easier to manage large classes and frequent requests from lecturers.

## Features

- Automatically converts lowercase subject codes to uppercase.
- Filters student data based on the provided subject code.
- Generates a professional-looking PDF class list that includes student names, matriculation numbers, and space for signatures.
- Adds alternating row colors for better readability in the PDF.
- Extracts course titles from a secondary Excel sheet to create a complete PDF heading.
- Automatically removes any existing PDF with the same filename to avoid confusion.

## Prerequisites

Make sure the following Python libraries are installed:

```bash
pip install pandas
pip install reportlab
