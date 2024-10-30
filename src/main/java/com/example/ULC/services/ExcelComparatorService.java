package com.example.ULC.services;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Service
public class ExcelComparatorService {

    public Map<String, List<String>> compareExcelFiles(MultipartFile file1, MultipartFile file2) throws IOException {
        // Хранит результаты сравнения
        Map<String, List<String>> result = new HashMap<>();

        // Чтение первого файла
        try (Workbook workbook1 = new XSSFWorkbook(file1.getInputStream())) {
            if (workbook1.getNumberOfSheets() == 0) {
                throw new IOException("В первом файле нет листов.");
            }

            // Проходим по всем листам первого файла
            for (int i = 0; i < workbook1.getNumberOfSheets(); i++) {
                Sheet sheet = workbook1.getSheetAt(i);
                // Проходим по всем ячейкам во второй строке (индекс 1)
                for (int col = 0; col < sheet.getRow(1).getPhysicalNumberOfCells(); col++) {
                    Row row = sheet.getRow(1);
                    Cell cell = row.getCell(col);

                    // Проверяем, что ячейка не пустая и содержит строку
                    // Проверяем, что ячейка не пустая
                    if (cell != null) {
                        String value = cellToString(cell); // Приводим ячейку к строке
                        List<String> columnValues = new ArrayList<>();

                        // Собираем значения с 22 по 35 ячейку (индексы 21-34)
                        for (int rowIndex = 21; rowIndex <= 34; rowIndex++) {
                            Row dataRow = sheet.getRow(rowIndex);
                            if (dataRow != null) {
                                Cell dataCell = dataRow.getCell(col);
                                if (dataCell != null && dataCell.getCellType() == CellType.STRING) {
                                    columnValues.add(dataCell.getStringCellValue());
                                }
                            }
                        }

                        // Сохраняем значения для ключа из второй строки
                        result.put(value, columnValues);
                    }
                }
            }
        }

        // Чтение второго файла и проверка на совпадения
        boolean matchesFound = false; // Переменная для отслеживания наличия совпадений

        try (Workbook workbook2 = new XSSFWorkbook(file2.getInputStream())) {
            if (workbook2.getNumberOfSheets() == 0) {
                throw new IOException("Во втором файле нет листов.");
            }

            // Проходим по всем листам второго файла
            for (int i = 0; i < workbook2.getNumberOfSheets(); i++) {
                Sheet sheet = workbook2.getSheetAt(i);
                // Проходим по всем ячейкам в первой строке (индекс 0)
                for (int col = 0; col < sheet.getRow(0).getPhysicalNumberOfCells(); col++) {
                    Row row = sheet.getRow(0);
                    Cell cell = row.getCell(col);

                    // Проверяем, что ячейка не пустая
                    if (cell != null) {
                        String keyValue = cellToString(cell); // Приводим ячейку к строке

                        // Проверяем каждое значение из первой строки первого файла
                        for (String value : result.keySet()) {
                            // Если найдено совпадение
                            if (value.equals(keyValue)) {
                                matchesFound = true; // Устанавливаем флаг, что совпадение найдено
                                List<String> columnValuesFromFile2 = new ArrayList<>();

                                // Собираем значения с 14 по 29 ячейку (индексы 13-28)
                                for (int rowIndex = 13; rowIndex <= 28; rowIndex++) {
                                    Row dataRow = sheet.getRow(rowIndex);
                                    if (dataRow != null) {
                                        Cell dataCell = dataRow.getCell(col);
                                        if (dataCell != null && dataCell.getCellType() == CellType.STRING) {
                                            columnValuesFromFile2.add(dataCell.getStringCellValue());
                                        }
                                    }
                                }

                                // Обновляем результат, добавляя значения из второго файла
                                result.put(value, List.of(result.get(value).toString(), columnValuesFromFile2.toString()));
                            }
                        }
                    }
                }
            }
        }

        // Если совпадения не найдены, добавляем соответствующее сообщение в результат
        if (!matchesFound) {
            result.put("Совпадения", List.of("Нет совпадений между файлами."));
        }

        return result; // Возвращаем результаты сравнения
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