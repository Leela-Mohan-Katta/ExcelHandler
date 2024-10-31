
# Excel Handler Spring Boot Application

## Overview

This Spring Boot application processes Excel files containing formulas and corresponding values from two sheets. It calculates the results based on the formulas provided and outputs the results in a new Excel file format. The application retrieves values from one sheet and applies formulas from another to generate a comprehensive output.

## Features

- Reads formulas from the first sheet (Sheet1) and corresponding values from the second sheet (Sheet2).
- Outputs a new Excel file with a specific format: `ID|NWA|TOTAL|1.0, 2.0, 3.0`.
- Allows users to access the functionality via a REST API endpoint.

## Prerequisites

- **Java 11** or higher.
- **Maven** for dependency management.
- **Apache POI** library for reading/writing Excel files.

## Setup Instructions

1. **Clone the Repository**:
   ```bash
   git clone <repository-url>
   cd <project-directory>
Add Dependencies: Ensure you have the following dependency in your pom.xml:

xml
Copy code
<dependency>
<groupId>org.apache.poi</groupId>
<artifactId>poi-ooxml</artifactId>
<version>5.2.2</version>
</dependency>
Run the Application:

Use your IDE or run the following command:
bash
Copy code
mvn spring-boot:run
Access the API: Open your web browser or use Postman to access the endpoint:

bash
Copy code
http://localhost:8080/excel/process?filePath=E:/Java/ACT.xlsx
Replace E:/Java/ACT.xlsx with the path to your Excel file.
Input Format
Sheet1: Contains formulas in the format NWA=A1+A2+A3.
Sheet2: Contains references and corresponding values in two columns:
Column A: Reference (e.g., A1, A2)
Column B: Value (e.g., 1.0, 2.0)
Output Format
The application outputs an Excel file with the following columns:
ID: The reference before the =.
NWA: The formula reference.
TOTAL: The string "TOTAL".
Values: The comma-separated values corresponding to the references used in the formula.
Example Output
An example output row might look like:

Copy code
911|NWA|TOTAL|1.0, 2.0, 3.0
Notes
Ensure that the provided file path to the Excel file is correct and accessible by the application.
The application handles basic errors related to file processing and will return error messages if the file cannot be processed.
Troubleshooting
If you encounter issues with the application, check the console logs for any error messages.
Ensure that the Excel file is not open in another program while trying to access it through this application