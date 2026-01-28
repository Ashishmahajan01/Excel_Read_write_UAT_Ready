package com.excelDemo.excel.service;

import com.excelDemo.excel.model.UserData;
import com.excelDemo.excel.repository.UserDataRepository;
import com.excelDemo.excel.util.ExcelIUtil;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

@Service
@RequiredArgsConstructor
public class ExcelDownloadService {

    private final UserDataRepository userDataRepository;

    public byte[] generateUsersExcel() throws IOException {
        List<UserData> users = userDataRepository.findAll();

        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            Sheet sheet = workbook.createSheet("Users Data");

            // Clear metadata
            ExcelIUtil.removeMetaData(workbook);

            // Create styles
            CellStyle headerStyle = createHeaderStyle(workbook);
            CellStyle dateStyle = createDateCellStyle(workbook);
            CellStyle numberStyle = createNumberCellStyle(workbook);

            // Create header row - ONLY THE COLUMNS THAT EXIST IN YOUR ENTITY
            Row headerRow = sheet.createRow(0);
            String[] headers = {
                    "ID", "Username", "Email", "Age", "Department",
                    "Salary", "Is Active", "Created At"
            };

            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerStyle);
            }

            // Create data rows
            int rowNum = 1;
            for (UserData user : users) {
                Row row = sheet.createRow(rowNum++);

                // Only include the fields that exist in your UserData entity
                createCell(row, 0, user.getId(), numberStyle);
                createCell(row, 1, user.getUsername(), null);
                createCell(row, 2, user.getEmail(), null);
                createCell(row, 3, user.getAge(), numberStyle);
                createCell(row, 4, user.getDepartment(), null);
                createCell(row, 5, user.getSalary(), numberStyle);
                createCell(row, 6, user.getIsActive(), null);

                // Created At with date formatting
                if (user.getCreatedAt() != null) {
                    Cell dateCell = row.createCell(7);
                    dateCell.setCellValue(user.getCreatedAt());
                    dateCell.setCellStyle(dateStyle);
                } else {
                    row.createCell(7).setCellValue("");
                }

            }

            for (int i = 0; i < headers.length; i++) {
                sheet.autoSizeColumn(i);
            }

            workbook.write(out);
            return out.toByteArray();
        }
    }

    private CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

    private CellStyle createDateCellStyle(Workbook workbook) {
        CellStyle dateStyle = workbook.createCellStyle();
        CreationHelper createHelper = workbook.getCreationHelper();
        dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy-mm-dd hh:mm:ss"));
        return dateStyle;
    }

    private CellStyle createNumberCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat("#,##0"));
        return style;
    }

    private void createCell(Row row, int columnIndex, Object value, CellStyle style) {
        if (value == null) {
            row.createCell(columnIndex).setCellValue("");
            return;
        }

        Cell cell = row.createCell(columnIndex);

        if (style != null) {
            cell.setCellStyle(style);
        }

        switch (value) {
            case String s -> cell.setCellValue(s);
            case Integer i -> cell.setCellValue(i);
            case Double v -> cell.setCellValue(v);
            case Boolean b -> cell.setCellValue(b);
            case Long l -> cell.setCellValue(l);
            default -> cell.setCellValue(value.toString());
        }
    }
}
