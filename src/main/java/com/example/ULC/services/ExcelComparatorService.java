package com.example.ULC.services;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

@Service
public class ExcelComparatorService {

    public Map<String, List<List<String>>> compareExcelFiles(File fileYT, File file1C) throws IOException {
        Map<String, List<String>> resultYT = new HashMap<>();
        Map<String, List<String>> result1C = new HashMap<>();
        Map<String, List<List<String>>> result = new HashMap<>();

        // Чтение первого файла - УТ
        try (Workbook workbookYT = new XSSFWorkbook(new FileInputStream(fileYT))) {
            if (workbookYT.getNumberOfSheets() == 0) {
                throw new IOException("В первом файле нет листов.");
            }

            for (int i = 0; i < workbookYT.getNumberOfSheets(); i++) {
                Sheet sheet = workbookYT.getSheetAt(i);
                Row headerRow = sheet.getRow(1); // Получаем вторую строку (индекс 1)

                if (headerRow == null) {
                    throw new IOException("Второй строки нет в листе " + (i + 1));
                }

                for (int col = 0; col < headerRow.getPhysicalNumberOfCells(); col++) {
                    Cell cell = headerRow.getCell(col);
                    if (cell != null) {
                        String value = cellToString(cell);
                        List<String> columnValues = new ArrayList<>();

                        // Собираем данные с 22 по 35 ячейку (индексы 21-34)
                        for (int rowIndex = 21; rowIndex <= 34; rowIndex++) {
                            Row dataRow = sheet.getRow(rowIndex);
                            if (dataRow != null) {
                                Cell dataCell = dataRow.getCell(col);
                                if (dataCell != null) {
                                    columnValues.add(cellToString(dataCell));
                                }
                            }
                        }

                        resultYT.put(value, columnValues); // Сохраняем пару value - список значений
                    }
                }
            }
        }

        // Чтение экспорта из 1С
        try (Workbook workbook1C = new XSSFWorkbook(new FileInputStream(file1C))) {
            if (workbook1C.getNumberOfSheets() == 0) {
                throw new IOException("В первом файле нет листов.");
            }

            for (int i = 0; i < workbook1C.getNumberOfSheets(); i++) {
                Sheet sheet = workbook1C.getSheetAt(i);
                Row headerRow = sheet.getRow(0); // Получаем вторую строку (индекс 1)

                if (headerRow == null) {
                    throw new IOException("Второй строки нет в листе " + (i + 1));
                }

                for (int col = 0; col < headerRow.getPhysicalNumberOfCells(); col++) {
                    Cell cell = headerRow.getCell(col);
                    if (cell != null) {
                        String value = cellToString(cell);
                        List<String> columnValues = new ArrayList<>();

                        // Собираем данные с 14 по 29 ячейку (индексы 13-28)
                        for (int rowIndex = 13; rowIndex <= 28; rowIndex++) {
                            Row dataRow = sheet.getRow(rowIndex);
                            if (dataRow != null) {
                                Cell dataCell = dataRow.getCell(col);
                                if (dataCell != null) {
                                    columnValues.add(cellToString(dataCell));
                                }
                            }
                        }

                        result1C.put(value, columnValues); // Сохраняем пару value - список значений
                    }
                }
            }
        }

        // Объединение карт
        for (String key : resultYT.keySet()) {
            if (result1C.containsKey(key)) { // Проверяем, есть ли ключ в обеих картах
                List<String> valuesFromYT = resultYT.get(key); // Получаем список из первой карты
                List<String> valuesFrom1C = result1C.get(key); // Получаем список из второй карты

                // Создаем новый список списков для результата
                List<List<String>> combinedValues = new ArrayList<>();
                combinedValues.add(valuesFromYT); // Добавляем первый список
                combinedValues.add(valuesFrom1C); // Добавляем второй список

                // Записываем в результирующую карту
                result.put(key, combinedValues);
            }
        }

        // Проход по всем ключам первой карты
        for (String key : resultYT.keySet()) {
            List<String> valuesFromYT = resultYT.get(key);
            List<String> valuesFrom1C = result1C.get(key);

            // Создаем множество для уникальных значений из resultYT
            Set<String> uniqueYT = new HashSet<>(valuesFromYT);

            // Создаем множество для уникальных значений из result1C
            Set<String> unique1C = new HashSet<>();
            if (valuesFrom1C != null) {
                unique1C.addAll(valuesFrom1C);
            }
            // Находим значения, которых нет во втором списке
            List<String> missingIn1C = new ArrayList<>(uniqueYT);
            missingIn1C.removeAll(unique1C); // Удаляем значения, которые есть в result1C

            // Находим значения, которых нет в первом списке
            List<String> missingInYT = new ArrayList<>(unique1C);
            missingInYT.removeAll(uniqueYT); // Удаляем значения, которые есть в resultYT

            // Создаем список списков для результата
            List<List<String>> combinedValues = new ArrayList<>();
            // Добавляем списки отсутствующих значений
            combinedValues.add(missingIn1C); // Отсутствующие значения из result1C
            combinedValues.add(missingInYT);  // Отсутствующие значения из resultYT

            // Проверяем, существует ли ключ в результирующей карте
            if (result.containsKey(key)) {
                // Если ключ существует, добавляем новые значения к существующим
                result.get(key).addAll(combinedValues);
            } else {
                // Если ключ не существует, добавляем новый
                result.put(key, combinedValues);
            }
        }

        return result;



//        // Чтение второго файла и проверка на совпадения
//        try (Workbook workbook2 = new XSSFWorkbook(file2.getInputStream())) {
//            if (workbook2.getNumberOfSheets() == 0) {
//                throw new IOException("Во втором файле нет листов.");
//            }
//
//            for (int i = 0; i < workbook2.getNumberOfSheets(); i++) {
//                Sheet sheet = workbook2.getSheetAt(i);
//                for (int col = 0; col < sheet.getRow(0).getPhysicalNumberOfCells(); col++) {
//                    Row row = sheet.getRow(0);
//                    Cell cell = row.getCell(col);
//
//                    if (cell != null) {
//                        String keyValue = cellToString(cell);
//                        List<List<String>> existingValues = result.get(keyValue);
//
//                        // Если совпадение найдено, добавляем данные из второго файла
//                        if (existingValues != null) {
//                            List<String> columnValuesFromFile2 = new ArrayList<>();
//
//                            for (int rowIndex = 13; rowIndex <= 28; rowIndex++) {
//                                Row dataRow = sheet.getRow(rowIndex);
//                                if (dataRow != null) {
//                                    Cell dataCell = dataRow.getCell(col);
//                                    if (dataCell != null) {
//                                        columnValuesFromFile2.add(cellToString(dataCell));
//                                    }
//                                }
//                            }
//
//                            // Обновляем все результаты с данными из второго файла
//                            for (List<String> values : existingValues) {
//                                values.addAll(columnValuesFromFile2); // Объединяем значения
//                            }
//                        }
//                    }
//                }
//            }
//        }

    }

    // Метод для приведения ячейки к строке
    private String cellToString(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "";
            default:
                return "Неподдерживаемый тип"; // Для неподдерживаемых типов
        }
    }
}