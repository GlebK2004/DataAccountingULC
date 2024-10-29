package com.example.ULC.services;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.HashSet;
import java.util.Set;

@Service
public class ExcelComparatorService {

    public Set<String> compareExcelFiles(MultipartFile file1, MultipartFile file2) throws IOException {
        Set<String> firstFileValues = new HashSet<>();
        Set<String> secondFileValues = new HashSet<>();

        // Чтение первого файла
        try (Workbook workbook1 = new XSSFWorkbook(file1.getInputStream())) {
            for (int i = 0; i < workbook1.getNumberOfSheets(); i++) {
                Sheet sheet = workbook1.getSheetAt(i);
                Row row = sheet.getRow(0); // Первая строка
                if (row != null) {
                    Cell cell = row.getCell(0); // Первая ячейка
                    if (cell != null) {
                        firstFileValues.add(cell.toString());
                    }
                }
            }
        }

        // Чтение второго файла
        try (Workbook workbook2 = new XSSFWorkbook(file2.getInputStream())) {
            for (int i = 0; i < workbook2.getNumberOfSheets(); i++) {
                Sheet sheet = workbook2.getSheetAt(i);
                Row row = sheet.getRow(0); // Первая строка
                if (row != null) {
                    Cell cell = row.getCell(0); // Первая ячейка
                    if (cell != null) {
                        secondFileValues.add(cell.toString());
                    }
                }
            }
        }

        // Поиск совпадений
        firstFileValues.retainAll(secondFileValues);
        return firstFileValues;
    }
}