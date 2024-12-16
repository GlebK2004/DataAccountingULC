package com.example.ULC.services;

import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;

@Service
public class DownloadService {


    public File downloadExcel(String sheetId)
    {
        try {
            // URL для экспорта таблицы
            String excelUrl = "https://docs.google.com/spreadsheets/d/" + sheetId + "/export?format=xlsx";

            // Устанавливаем соединение
            URL url = new URL(excelUrl);
            HttpURLConnection connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("GET");

            // Проверяем код ответа
            int responseCode = connection.getResponseCode();
            if (responseCode == HttpURLConnection.HTTP_OK) {
                // Скачиваем и сохраняем файл
                File file = new File("output.xlsx");
                try (InputStream inputStream = connection.getInputStream();
                     FileOutputStream fileOutputStream = new FileOutputStream(file)) {

                    byte[] buffer = new byte[4096];
                    int bytesRead;
                    while ((bytesRead = inputStream.read(buffer)) != -1) {
                        fileOutputStream.write(buffer, 0, bytesRead);
                    }
                }
                return file;
            } else {
                System.out.println("Ошибка при загрузке файла. Код ответа: " + responseCode);
                return null;
            }
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }


    }


}
