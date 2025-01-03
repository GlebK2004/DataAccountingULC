package com.example.ULC.controllers;

import com.example.ULC.services.DownloadService;
import com.example.ULC.services.ExcelComparatorService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;

@Controller
public class ExcelController {
    static final String ID_YT = "1rZOYBNWty6BTm5OFDuaRsWmdTW7QFjKHIyZpTuqx4nE";
    static final String ID_1C = "1Qbh5rSJ9Xbq6-Da2gC4qQNTBgh3j97h7zracHrIgxSA";
    @Autowired
    private ExcelComparatorService excelComparatorService;
    @Autowired
    private DownloadService downloadService;

//    @PostMapping("/compare")
//    public ResponseEntity<Map<String, List<List<String>>>> compareExcelFiles(@RequestParam("file1") MultipartFile file1,
//                                                                       @RequestParam("file2") MultipartFile file2) {
//        try {
//            Map<String, List<List<String>>> matches = excelComparatorService.compareExcelFiles(file1, file2);
//            return ResponseEntity.ok(matches);
//        } catch (IOException e) {
//            return ResponseEntity.badRequest().body(null);
//        }
//    }
    @GetMapping("/")
    public String mainPage(Model model)
    {
        return "main_page";
    }

    @GetMapping("/groups")
    public String downloadFiles(Model model) throws IOException {
       File fileYT = downloadService.downloadExcel(ID_YT);
       File file1C = downloadService.downloadExcel(ID_1C);
       model.addAttribute("groupsList", excelComparatorService.compareExcelFiles(fileYT, file1C));
       return "groups";
    }

}