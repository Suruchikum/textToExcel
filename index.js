
const fs = require("fs");
const path = require("path");
const express = require("express");
const ExcelJS = require("exceljs");

const app = express();
const port = 3000;

app.get("/home", (req, res) => {
  res.send("hello home");
});

app.get("/employees", (req, res) => {
  fs.readFile("data.txt", "utf-8", (err, data) => {
    if (err) {
      return res.status(500).send("Error reading file");
    }

    const employees = [];
    const lines = data.split("\n");

    let currentEmployee = null;

    lines.forEach((line) => {
      const nameMatch = line.match(/Name\s+:\s+(.+?)\s+\|/);
      if (nameMatch) {
        if (currentEmployee) {
          // Check for duplicates before pushing
          if (!employees.find(emp => emp.name === currentEmployee.name && emp.id === currentEmployee.id && emp.totalSalary === currentEmployee.totalSalary)) {
            employees.push(currentEmployee);
          }
        }
        currentEmployee = { name: nameMatch[1] };
      }

      const idMatch = line.match(/Id\s+:\s+(\d+)\s+\|/);
      if (idMatch) {
        currentEmployee.id = parseInt(idMatch[1], 10);
      }

      const salaryMatch = line.match(/Take Home Pay\s+(\d+,\d+\.\d{2})/);
      if (salaryMatch) {
        currentEmployee.totalSalary = parseFloat(salaryMatch[1].replace(",", ""));
      }
    });

    if (currentEmployee) {
      // Check for duplicates before pushing
      if (!employees.find(emp => emp.name === currentEmployee.name && emp.id === currentEmployee.id && emp.totalSalary === currentEmployee.totalSalary)) {
        employees.push(currentEmployee);
      }
    }

    console.log(`Parsed employees: ${JSON.stringify(employees, null, 2)}`);

    
    const outputFilePath = path.join(__dirname, "employees.json");
    fs.writeFile(outputFilePath, JSON.stringify(employees, null, 2), (err) => {
      if (err) {
        return res.status(500).send("Error writing JSON file");
      }

    
      createWorkbook(employees, (err) => {
        if (err) {
          return res.status(500).send("Error creating Excel file");
        }
        res.json({ message: "Data written to employees.json and employees.xlsx", employees });
      });
    });
  });
});

function createWorkbook(employees, callback) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Employees');

  // Add headers
  worksheet.getRow(1).values = ["Name", "Id", "Take Home Pay"];

  // Add employee data
  employees.forEach((emp) => {
    worksheet.addRow([emp.name, emp.id, emp.totalSalary || 0.00]);
  });

  // Save the workbook to a file
  const filePath = path.join(__dirname, 'employees.xlsx');
  workbook.xlsx.writeFile(filePath)
    .then(() => {
      console.log(`Workbook saved to ${filePath}`);
      callback();
    })
    .catch(err => {
      console.error('Error writing workbook:', err);
      callback(err);
    });
}

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
