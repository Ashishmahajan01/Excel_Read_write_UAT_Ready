package com.excelDemo.excel.controller;

import com.excelDemo.excel.dto.FileResponse;
import com.excelDemo.excel.exception.ExcelProcessingException;
import com.excelDemo.excel.exception.FileUploadException;
import com.excelDemo.excel.exception.InvalidFileFormatException;
import com.excelDemo.excel.model.UserData;
import com.excelDemo.excel.service.ExcelDownloadService;
import com.excelDemo.excel.service.ExcelUploadService;
import com.excelDemo.excel.util.ExcelIUtil;
import lombok.RequiredArgsConstructor;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.time.LocalDateTime;
import java.util.List;

@RestController
@RequestMapping("/api/excel")
@RequiredArgsConstructor
public class ExcelUploadController {


    private final ExcelUploadService excelUploadService;
    private final ExcelDownloadService excelDownloadService;

    @PostMapping("/upload")
    public ResponseEntity<FileResponse> uploadExcelFile(@RequestParam("file") MultipartFile file) {
        try {
            // Validate file format using Content-Type and magic bytes (file signature)
            if (!ExcelIUtil.hasExcelFormat(file)) {
                return ResponseEntity.status(400).body(new FileResponse(
                        file.getOriginalFilename(),
                        (int) file.getSize(),
                        LocalDateTime.now(),
                        "Failed",
                        List.of("Invalid file format. Please upload a valid Excel file (.xlsx). File signature validation failed."),
                        0
                ));
            }

            FileResponse response = excelUploadService.processExcelFile(file);

            if (response.getErrors() != null && !response.getErrors().isEmpty()) {
                return ResponseEntity.badRequest().body(response);
            }

            return ResponseEntity.status(200).body(response);

        } catch (InvalidFileFormatException | ExcelProcessingException e) {
            // These will be handled by GlobalExceptionHandler
            throw e;
        } catch (com.excelDemo.excel.exception.ValidationException e) {
            // Validation exceptions return 400 Bad Request
            throw e;
        } catch (Exception e) {
            throw new FileUploadException("Failed to upload file: " + e.getMessage(), e);
        }
    }

    @GetMapping("/users")
    public ResponseEntity<byte[]> getAllUsersAsExcel() {
        try {
            byte[] excelReport = excelDownloadService.generateUsersExcel();

            return new ResponseEntity<>(
                    excelReport,
                    ExcelIUtil.createExcelHeaders("User_Demo.xlsx", excelReport),
                    HttpStatus.OK
            );
        } catch (Exception e) {
            throw new RuntimeException("Error generating Excel file: " + e.getMessage(), e);
        }
    }
}
