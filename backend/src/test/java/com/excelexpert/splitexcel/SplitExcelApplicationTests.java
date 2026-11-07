package com.excelexpert.splitexcel;

import com.excelexpert.splitexcel.service.ExcelSplitService;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

@SpringBootTest
class SplitExcelApplicationTests {

    @Autowired
    private ExcelSplitService excelSplitService;

    @Test
    void contextLoads() {
        assertNotNull(excelSplitService);
    }

    @Test
    void testHssfSplitPreservesFormatting() throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try (HSSFWorkbook wb = new HSSFWorkbook()) {
            Sheet s1 = wb.createSheet("Table1");
            Sheet s2 = wb.createSheet("Table2");

            Row r = s1.createRow(0);
            Cell c = r.createCell(0);
            c.setCellValue("Hello Khmer");

            CellStyle style = wb.createCellStyle();
            Font font = wb.createFont();
            font.setFontName("Khmer OS Battambang");
            font.setBold(true);
            style.setFont(font);
            style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setBorderBottom(BorderStyle.THICK);
            c.setCellStyle(style);

            wb.write(baos);
        }

        byte[] inputBytes = baos.toByteArray();
        byte[] outputBytes = excelSplitService.splitSingleSheet(inputBytes, "Table1", null);

        // Verify output is valid HSSF workbook
        try (Workbook outWb = WorkbookFactory.create(new ByteArrayInputStream(outputBytes))) {
            assertEquals(1, outWb.getNumberOfSheets());
            assertEquals("Table1", outWb.getSheetName(0));

            Sheet outSheet = outWb.getSheetAt(0);
            Cell outCell = outSheet.getRow(0).getCell(0);
            assertEquals("Hello Khmer", outCell.getStringCellValue());

            CellStyle outStyle = outCell.getCellStyle();
            Font outFont = outWb.getFontAt(outStyle.getFontIndex());
            assertEquals("Khmer OS Battambang", outFont.getFontName());
            assertTrue(outFont.getBold());
            assertEquals(BorderStyle.THICK, outStyle.getBorderBottom());
        }
    }

    @Test
    void testXssfSplitPreservesFormatting() throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet s1 = wb.createSheet("DataSheet1");
            Sheet s2 = wb.createSheet("DataSheet2");

            Row r = s1.createRow(0);
            Cell c = r.createCell(0);
            c.setCellValue("Styled Text");

            CellStyle style = wb.createCellStyle();
            Font font = wb.createFont();
            font.setFontName("Khmer OS Battambang");
            font.setBold(true);
            style.setFont(font);
            style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setBorderBottom(BorderStyle.DOUBLE);
            c.setCellStyle(style);

            wb.write(baos);
        }

        byte[] inputBytes = baos.toByteArray();
        byte[] outputBytes = excelSplitService.splitSingleSheet(inputBytes, "DataSheet1", null);

        // Verify output is valid XSSF workbook
        try (Workbook outWb = WorkbookFactory.create(new ByteArrayInputStream(outputBytes))) {
            assertEquals(1, outWb.getNumberOfSheets());
            assertEquals("DataSheet1", outWb.getSheetName(0));

            Sheet outSheet = outWb.getSheetAt(0);
            Cell outCell = outSheet.getRow(0).getCell(0);
            assertEquals("Styled Text", outCell.getStringCellValue());

            CellStyle outStyle = outCell.getCellStyle();
            Font outFont = outWb.getFontAt(outStyle.getFontIndex());
            assertEquals("Khmer OS Battambang", outFont.getFontName());
            assertTrue(outFont.getBold());
            assertEquals(BorderStyle.DOUBLE, outStyle.getBorderBottom());
        }
    }

    @Test
    void testHiddenSheetsAreExcluded() throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet visibleSheet1 = wb.createSheet("Visible_1");
            visibleSheet1.createRow(0).createCell(0).setCellValue("Data 1");

            Sheet hiddenSheet = wb.createSheet("Hidden_Sheet");
            hiddenSheet.createRow(0).createCell(0).setCellValue("Confidential Data");
            wb.setSheetVisibility(1, SheetVisibility.HIDDEN);

            Sheet veryHiddenSheet = wb.createSheet("Very_Hidden_Sheet");
            veryHiddenSheet.createRow(0).createCell(0).setCellValue("Secret Formula");
            wb.setSheetVisibility(2, SheetVisibility.VERY_HIDDEN);

            Sheet visibleSheet2 = wb.createSheet("Visible_2");
            visibleSheet2.createRow(0).createCell(0).setCellValue("Data 2");

            wb.write(baos);
        }

        byte[] inputBytes = baos.toByteArray();

        // 1. Inspect workbook must only report visible sheets
        var meta = excelSplitService.inspectWorkbook(inputBytes, "Master.xlsx");
        assertEquals(2, meta.sheets().size());
        assertEquals("Visible_1", meta.sheets().get(0).name());
        assertEquals("Visible_2", meta.sheets().get(1).name());

        // 2. Packaging ZIP must only include visible sheets
        byte[] zipBytes = excelSplitService.splitAndPackageZip(
                inputBytes,
                "Master.xlsx",
                new com.excelexpert.splitexcel.model.SplitConfigRequest(
                        "OutputFolder",
                        "",
                        "_",
                        "original",
                        1,
                        "2",
                        null,
                        null,
                        false
                )
        );

        try (java.util.zip.ZipInputStream zis = new java.util.zip.ZipInputStream(new ByteArrayInputStream(zipBytes))) {
            java.util.List<String> entryNames = new java.util.ArrayList<>();
            java.util.zip.ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                entryNames.add(entry.getName());
            }

            assertEquals(2, entryNames.size());
            assertTrue(entryNames.contains("OutputFolder/Visible_1.xlsx"));
            assertTrue(entryNames.contains("OutputFolder/Visible_2.xlsx"));
            assertFalse(entryNames.stream().anyMatch(n -> n.contains("Hidden")));
        }
    }

    @Test
    void testFirstSheetIsInfoNaming() throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("Menu_Info").createRow(0).createCell(0).setCellValue("Table of contents");
            wb.createSheet("Details_A").createRow(0).createCell(0).setCellValue("Detail A");
            wb.createSheet("Details_B").createRow(0).createCell(0).setCellValue("Detail B");
            wb.write(baos);
        }

        byte[] inputBytes = baos.toByteArray();
        byte[] zipBytes = excelSplitService.splitAndPackageZip(
                inputBytes,
                "Book.xlsx",
                new com.excelexpert.splitexcel.model.SplitConfigRequest(
                        "Book_Folder",
                        "Report",
                        "_",
                        "sequence",
                        1,
                        "2",
                        null,
                        null,
                        true
                )
        );

        try (java.util.zip.ZipInputStream zis = new java.util.zip.ZipInputStream(new ByteArrayInputStream(zipBytes))) {
            java.util.List<String> entryNames = new java.util.ArrayList<>();
            java.util.zip.ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                entryNames.add(entry.getName());
            }

            assertEquals(3, entryNames.size());
            // First sheet is Info -> Prefix only
            assertTrue(entryNames.contains("Book_Folder/Report.xlsx"));
            // Subsequent sheets start sequence from startNumber (1 -> 01, 02)
            assertTrue(entryNames.contains("Book_Folder/Report_01.xlsx"));
            assertTrue(entryNames.contains("Book_Folder/Report_02.xlsx"));
        }
    }

    @Test
    void testCustomInfoFileName() throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try (Workbook wb = new XSSFWorkbook()) {
            wb.createSheet("Sheet_Menu").createRow(0).createCell(0).setCellValue("Menu");
            wb.createSheet("Sheet_Data1").createRow(0).createCell(0).setCellValue("Data 1");
            wb.createSheet("Sheet_Data2").createRow(0).createCell(0).setCellValue("Data 2");
            wb.write(baos);
        }

        byte[] inputBytes = baos.toByteArray();
        byte[] zipBytes = excelSplitService.splitAndPackageZip(
                inputBytes,
                "Book.xlsx",
                new com.excelexpert.splitexcel.model.SplitConfigRequest(
                        "Book_Folder",
                        "Report",
                        "_",
                        "sequence",
                        1,
                        "2",
                        null,
                        null,
                        true,
                        "Index_Menu"
                )
        );

        try (java.util.zip.ZipInputStream zis = new java.util.zip.ZipInputStream(new ByteArrayInputStream(zipBytes))) {
            java.util.List<String> entryNames = new java.util.ArrayList<>();
            java.util.zip.ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                entryNames.add(entry.getName());
            }

            assertEquals(3, entryNames.size());
            // First sheet is Info with custom name -> Index_Menu.xlsx
            assertTrue(entryNames.contains("Book_Folder/Index_Menu.xlsx"));
            // Subsequent sheets start sequence from startNumber (1 -> 01, 02)
            assertTrue(entryNames.contains("Book_Folder/Report_01.xlsx"));
            assertTrue(entryNames.contains("Book_Folder/Report_02.xlsx"));
        }
    }
}

