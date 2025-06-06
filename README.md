# Excel Data-Driven Automation Testing (Selenium + Apache POI)

This project automates the testing of a **Certificate of Deposit (CD) Calculator** on the CIT Bank website using **Java**, **Selenium WebDriver**, and **Excel integration via Apache POI**. The automation script reads multiple test inputs from an Excel file and validates the output shown on the web application.

---

## 🔧 Tools & Technologies Used
- Java
- Selenium WebDriver
- Apache POI (for reading/writing Excel files)
- Eclipse IDE
- ChromeDriver

---

## 📄 Project Description

- The script opens the CD calculator page and enters values like **deposit amount**, **interest rate**, **duration (months)**, and **compounding frequency** from an Excel file.
- It clicks the **Calculate** button, captures the result from the web app, and compares it with the **expected value** in the Excel file.
- After each test, it writes **"Passed"** or **"Failed"** back to the Excel file and applies a green or red cell background using Apache POI.

---

## 📁 Project Structure

excelproject/
│
├── Excelproject.java # Main test logic using Selenium
├── ExcelUtil.java # Utility class for Excel read/write operations
├── Book1.xlsx # Test data (input and expected result)
└── README.md # Project documentation

---

## 🚀 How to Run the Project

1. Open the project in **Eclipse IDE**.
2. Add the required external JARs:
   - Selenium WebDriver
   - Apache POI (poi + poi-ooxml)
3. Make sure the Excel file path in the code is correct:
	String fpath="C:\\Users\\user\\eclipse-workspace\\sele\\src\\test\\java\\TestData\\Book1.xlsx";
4. Run the `Excelproject.java` file as a Java application.
5. Check the Excel file for updated result statuses (Pass/Fail).

---

## ✅ Features

- Data-driven automation without any test framework
- Dynamic interaction with web elements (input fields, dropdowns)
- Excel cell color formatting (green/red) based on test results
- Pop-up handling for cookie banner
- Clean code separation via utility class for Excel operations

---

## 🌐 Tested Site

- [CIT Bank CD Calculator](https://www.cit.com/cit-bank/resources/calculators/certificate-of-deposit-calculator)

---

## 📌 Notes

- No TestNG or JUnit used—tests are run using plain Java logic
- Apache POI is used both for reading inputs and writing test results
- Script uses implicit and explicit waits to ensure element stability

