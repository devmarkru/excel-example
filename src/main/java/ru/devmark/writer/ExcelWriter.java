package ru.devmark.writer;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.List;

public class ExcelWriter {

    private static final List<ClientInfo> CLIENTS = Arrays.asList(
            new ClientInfo("Иванов Ваня", LocalDate.of(1988, 1, 20)),
            new ClientInfo("Петров Петя", LocalDate.of(1992, 10, 6))
    );

    public void write(String filename) throws IOException {
        var workbook = new XSSFWorkbook();
        var sheet = createSheet(workbook);
        createHeader(workbook, sheet);
        createCells(workbook, sheet);
        try (var outputStream = new FileOutputStream(filename)) {
            workbook.write(outputStream);
        }
        workbook.close();
    }

    private Sheet createSheet(XSSFWorkbook workbook) {
        var sheet = workbook.createSheet("Клиенты");
        sheet.setColumnWidth(0, 4000);
        sheet.setColumnWidth(1, 6000);
        sheet.setColumnWidth(2, 4000);
        return sheet;
    }

    private void createHeader(XSSFWorkbook workbook, Sheet sheet) {
        var header = sheet.createRow(0);

        var headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        var font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 14);
        font.setBold(true);
        headerStyle.setFont(font);

        var headerCell = header.createCell(0);
        headerCell.setCellValue("Имя");
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(1);
        headerCell.setCellValue("Дата рождения");
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(2);
        headerCell.setCellValue("Возраст");
        headerCell.setCellStyle(headerStyle);
    }

    private void createCells(XSSFWorkbook workbook, Sheet sheet) {
        var style = workbook.createCellStyle();
        style.setWrapText(true);
        var createHelper = workbook.getCreationHelper();
        style.setDataFormat(createHelper.createDataFormat().getFormat("dd.mm.yyyy"));
        sheet.setDefaultColumnStyle(1, style);

        for (var i = 0; i < CLIENTS.size(); i++) {
            var client = CLIENTS.get(i);
            var row = sheet.createRow(i + 1);

            var cell = row.createCell(0);
            cell.setCellValue(client.getName());

            cell = row.createCell(1);
            cell.setCellValue(client.getBirthDate());

            cell = row.createCell(2);
            cell.setCellFormula(String.format("TEXT(DATEDIF(B%s,NOW(),\"Y\"), \"0\")", i + 2));
        }
    }
}
