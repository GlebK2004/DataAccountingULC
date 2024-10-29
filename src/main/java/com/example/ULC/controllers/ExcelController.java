package com.example.ULC.controllers;

import com.example.ULC.services.ExcelComparatorService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.Set;

@RestController
@RequestMapping("/api/excel")
public class ExcelController {

    @Autowired
    private ExcelComparatorService excelComparatorService;

    @PostMapping("/compare")
    public ResponseEntity<Set<String>> compareExcelFiles(@RequestParam("file1") MultipartFile file1,
                                                         @RequestParam("file2") MultipartFile file2) {
        try {
            Set<String> matches = excelComparatorService.compareExcelFiles(file1, file2);
            return ResponseEntity.ok(matches);
        } catch (IOException e) {
            return ResponseEntity.badRequest().body(null);
        }
    }
}