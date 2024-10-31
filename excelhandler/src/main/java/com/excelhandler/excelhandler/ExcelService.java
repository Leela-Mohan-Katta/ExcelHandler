package com.excelhandler.excelhandler;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Service
public class ExcelService {

    public Map<String, String> processExcel(String filePath) throws IOException {
        Map<String, String> formulaResults = new HashMap<>();

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet1 = workbook.getSheetAt(0); // Sheet1 for formulas
            Sheet sheet2 = workbook.getSheetAt(1); // Sheet2 for values

            // Read values from Sheet2 and store in a map for quick lookup
            Map<String, Integer> valueMap = new HashMap<>();
            for (Row row : sheet2) {
                Cell referenceCell = row.getCell(0); // Cell like A1, A2
                Cell valueCell = row.getCell(1);     // Corresponding value

                if (referenceCell != null && valueCell != null) {
                    String reference = referenceCell.getStringCellValue();
                    int value = (int) valueCell.getNumericCellValue(); // Cast to integer
                    valueMap.put(reference, value);
                }
            }

            // Process formulas from Sheet1
            for (Row row : sheet1) {
                Cell idCell = row.getCell(0);       // ID cell
                Cell formulaCell = row.getCell(1);   // Formula cell, e.g., NWA=A1+A2+...

                if (idCell != null && formulaCell != null) {
                    int id = (int) idCell.getNumericCellValue();
                    String formula = formulaCell.getStringCellValue();
                    String values = getValuesUsed(formula, valueMap);
                    String result = id + "|" + formula.split("=")[0] + "|TOTAL|" + values;
                    formulaResults.put(formula, result);
                }
            }
        }

        return formulaResults;
    }

    private String getValuesUsed(String formula, Map<String, Integer> valueMap) {
        String[] tokens = formula.split("=");
        if (tokens.length < 2) return "";
        String[] references = tokens[1].split("\\+");

        List<String> valuesList = new ArrayList<>();
        for (String ref : references) {
            Integer value = valueMap.get(ref.trim());
            if (value != null) {
                valuesList.add(String.valueOf(value)); // Add value as a string
            }
        }
        // Join values with ", " and return
        return String.join(",", valuesList);
    }
}
