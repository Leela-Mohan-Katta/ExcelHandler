package com.excelhandler.excelhandler;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.bind.annotation.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

@RestController
@RequestMapping("/excel")
public class ExcelController {

    private final ExcelService excelService;

    public ExcelController(ExcelService excelService) {
        this.excelService = excelService;
    }

    @GetMapping("/process")
    public Map<String, String> processExcel(@RequestParam String filePath) {
        try {
            return excelService.processExcel(filePath);
        } catch (IOException e) {
            e.printStackTrace();
            return Map.of("error", "Could not process the file: " + e.getMessage());
        }
    }
}