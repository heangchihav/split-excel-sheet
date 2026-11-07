package com.excelexpert.splitexcel.service;

import com.excelexpert.splitexcel.model.SheetMetaDto;
import com.excelexpert.splitexcel.model.SplitConfigRequest;
import com.excelexpert.splitexcel.model.WorkbookMetaDto;
import org.apache.commons.io.output.ByteArrayOutputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

@Service
public class ExcelSplitService {

    private static final Logger log = LoggerFactory.getLogger(ExcelSplitService.class);

    /**
     * Inspects workbook metadata, sheets, row/col counts, and preview data.
     */
    public WorkbookMetaDto inspectWorkbook(byte[] fileBytes, String fileName) throws IOException {
        boolean isBinaryXls = isBinaryXlsFormat(fileBytes);
        DataFormatter dataFormatter = new DataFormatter();

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(fileBytes))) {
            int sheetCount = workbook.getNumberOfSheets();
            List<SheetMetaDto> sheets = new ArrayList<>();

            for (int i = 0; i < sheetCount; i++) {
                // If sheet is hidden or very hidden in master Excel, do NOT include it
                if (workbook.isSheetHidden(i) || workbook.isSheetVeryHidden(i)) {
                    log.info("Excluding hidden sheet from split: index={}, name='{}'", i, workbook.getSheetName(i));
                    continue;
                }

                Sheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName();
                int rowCount = sheet.getPhysicalNumberOfRows();
                int colCount = 0;

                List<String> headers = new ArrayList<>();
                List<List<String>> sampleRows = new ArrayList<>();

                int firstRowNum = sheet.getFirstRowNum();
                int lastRowNum = sheet.getLastRowNum();

                if (firstRowNum >= 0 && lastRowNum >= firstRowNum) {
                    // Sample up to 6 rows (1 header + 5 data rows)
                    int rowsToSample = Math.min(6, lastRowNum - firstRowNum + 1);
                    int maxColsFound = 0;

                    for (int r = firstRowNum; r < firstRowNum + rowsToSample; r++) {
                        Row row = sheet.getRow(r);
                        List<String> rowVals = new ArrayList<>();

                        if (row != null) {
                            short lastCellNum = row.getLastCellNum();
                            if (lastCellNum > maxColsFound) {
                                maxColsFound = lastCellNum;
                            }
                            int colsToRead = Math.min(8, (int) lastCellNum);

                            for (int c = 0; c < colsToRead; c++) {
                                Cell cell = row.getCell(c);
                                String cellVal = dataFormatter.formatCellValue(cell);
                                rowVals.add(cellVal != null ? cellVal : "");
                            }
                        }

                        if (r == firstRowNum) {
                            headers = rowVals;
                        } else {
                            sampleRows.add(rowVals);
                        }
                    }
                    colCount = maxColsFound;
                }

                sheets.add(new SheetMetaDto(i, sheetName, rowCount, colCount, headers, sampleRows));
            }

            return new WorkbookMetaDto(fileName, fileBytes.length, isBinaryXls, sheets);
        }
    }

    /**
     * Splits out a single sheet into an independent workbook preserving 100% of formatting and fonts.
     */
    public byte[] splitSingleSheet(byte[] fileBytes, String targetSheetName, String customTabName) throws IOException {
        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(fileBytes))) {
            int targetIndex = findSheetIndex(workbook, targetSheetName);
            if (targetIndex == -1) {
                throw new IllegalArgumentException("Sheet '" + targetSheetName + "' not found in workbook.");
            }

            // Remove all sheets other than targetIndex (loop backwards to preserve indices)
            for (int i = workbook.getNumberOfSheets() - 1; i >= 0; i--) {
                if (i != targetIndex) {
                    workbook.removeSheetAt(i);
                }
            }

            // Rename the remaining sheet if customTabName is provided
            if (customTabName != null && !customTabName.trim().isEmpty()) {
                String safeTabName = sanitizeSheetTabName(customTabName.trim());
                workbook.setSheetName(0, safeTabName);
            }

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            byte[] outBytes = baos.toByteArray();

            // Sanitize any Devanagari font fallbacks if it's an OOXML (.xlsx) package
            if (isZipPackage(outBytes)) {
                return sanitizeDevanagariFonts(outBytes);
            }
            return outBytes;
        }
    }

    /**
     * Splits multiple sheets from a workbook and bundles them into a ZIP archive inside a dedicated master folder.
     */
    public byte[] splitAndPackageZip(byte[] fileBytes, String originalFileName, SplitConfigRequest config) throws IOException {
        String baseName = originalFileName.replaceAll("\\.[^.]+$", "");
        String rawFolderName = config.folderName() != null && !config.folderName().trim().isEmpty()
                ? config.folderName().trim()
                : baseName;
        String folderName = sanitizeFileName(rawFolderName, "Excel_Sheets");

        WorkbookMetaDto meta = inspectWorkbook(fileBytes, originalFileName);
        boolean isXls = isBinaryXlsFormat(fileBytes);
        String fileExt = isXls ? "xls" : "xlsx";

        Set<String> visibleSheetNames = new HashSet<>(meta.sheets().stream().map(SheetMetaDto::name).toList());
        List<String> selected = (config.selectedSheets() != null && !config.selectedSheets().isEmpty()
                ? config.selectedSheets()
                : meta.sheets().stream().map(SheetMetaDto::name).toList())
                .stream()
                .filter(visibleSheetNames::contains)
                .toList();

        ByteArrayOutputStream zipBaos = new ByteArrayOutputStream();
        try (ZipOutputStream zos = new ZipOutputStream(zipBaos)) {
            Set<String> usedFileNames = new HashSet<>();
            boolean hasInfoSheet = config.resolveFirstSheetIsInfo() && !selected.isEmpty();

            for (int idx = 0; idx < selected.size(); idx++) {
                String originalSheetName = selected.get(idx);
                String customName = config.customSheetNames() != null ? config.customSheetNames().get(originalSheetName) : null;
                boolean isInfo = hasInfoSheet && idx == 0;
                int effectiveIndex = (hasInfoSheet && idx > 0) ? idx - 1 : idx;
                int totalSeqSheets = hasInfoSheet ? Math.max(1, selected.size() - 1) : selected.size();

                String sheetFileName = resolveSheetFileName(
                        config.resolvePrefix(),
                        config.resolveSeparator(),
                        config.resolveNamingMode(),
                        originalSheetName,
                        customName,
                        effectiveIndex,
                        config.resolveStartNumber(),
                        config.resolveZeroPadding(),
                        totalSeqSheets,
                        fileExt,
                        isInfo,
                        config.resolveInfoFileName()
                );

                String uniqueFileName = getUniqueFileName(sheetFileName, usedFileNames);
                byte[] singleSheetBytes = splitSingleSheet(fileBytes, originalSheetName, customName);

                ZipEntry entry = new ZipEntry(folderName + "/" + uniqueFileName);
                zos.putNextEntry(entry);
                zos.write(singleSheetBytes);
                zos.closeEntry();
            }
        }

        return zipBaos.toByteArray();
    }

    /**
     * Formats sheet file name according to user configuration and matching file extension.
     */
    public String resolveSheetFileName(
            String prefix,
            String separator,
            String namingMode,
            String originalSheetName,
            String customName,
            int index,
            int startNumber,
            String zeroPadding,
            int totalSheets,
            String extension,
            boolean isInfoSheet,
            String infoFileName
    ) {
        String cleanExt = extension != null ? extension.replaceAll("^\\.", "") : "xlsx";
        if (customName != null && !customName.trim().isEmpty()) {
            String clean = sanitizeFileName(customName.trim(), "Sheet");
            return clean.toLowerCase().endsWith("." + cleanExt) ? clean : clean + "." + cleanExt;
        }

        String cleanPrefix = prefix != null ? sanitizeFileName(prefix.trim(), "") : "";
        String cleanInfoName = infoFileName != null ? sanitizeFileName(infoFileName.trim(), "") : "";

        // If this is the Info sheet, use custom info file name if given, else prefix, else "Info"
        if (isInfoSheet) {
            String infoBase = !cleanInfoName.isEmpty() ? cleanInfoName : (!cleanPrefix.isEmpty() ? cleanPrefix : "Info");
            return infoBase + "." + cleanExt;
        }

        String sep = !cleanPrefix.isEmpty() ? separator : "";
        int currentNum = startNumber + index;
        String numStr = formatSequenceNumber(currentNum, zeroPadding, totalSheets);
        String cleanOriginalName = sanitizeFileName(originalSheetName, "Sheet");

        String body;
        switch (namingMode) {
            case "sequence" -> body = numStr;
            case "sequence_original" -> body = numStr + "_" + cleanOriginalName;
            case "original" -> body = cleanOriginalName;
            default -> body = cleanOriginalName;
        }

        String baseName = !cleanPrefix.isEmpty() ? cleanPrefix + sep + body : body;
        return baseName + "." + cleanExt;
    }

    public String resolveSheetFileName(
            String prefix,
            String separator,
            String namingMode,
            String originalSheetName,
            String customName,
            int index,
            int startNumber,
            String zeroPadding,
            int totalSheets,
            String extension,
            boolean isInfoSheet
    ) {
        return resolveSheetFileName(prefix, separator, namingMode, originalSheetName, customName, index, startNumber, zeroPadding, totalSheets, extension, isInfoSheet, null);
    }

    public String resolveSheetFileName(
            String prefix,
            String separator,
            String namingMode,
            String originalSheetName,
            String customName,
            int index,
            int startNumber,
            String zeroPadding,
            int totalSheets,
            String extension
    ) {
        return resolveSheetFileName(prefix, separator, namingMode, originalSheetName, customName, index, startNumber, zeroPadding, totalSheets, extension, false);
    }

    public String resolveSheetFileName(
            String prefix,
            String separator,
            String namingMode,
            String originalSheetName,
            String customName,
            int index,
            int startNumber,
            String zeroPadding,
            int totalSheets
    ) {
        return resolveSheetFileName(prefix, separator, namingMode, originalSheetName, customName, index, startNumber, zeroPadding, totalSheets, "xlsx", false);
    }

    private String formatSequenceNumber(int num, String padding, int totalItems) {
        return switch (padding) {
            case "none" -> String.valueOf(num);
            case "2" -> String.format("%02d", num);
            case "3" -> String.format("%03d", num);
            case "auto" -> {
                int digitsNeeded = Math.max(2, String.valueOf(totalItems).length());
                yield String.format("%0" + digitsNeeded + "d", num);
            }
            default -> String.valueOf(num);
        };
    }

    private int findSheetIndex(Workbook workbook, String targetName) {
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            String name = workbook.getSheetName(i);
            if (name.equalsIgnoreCase(targetName) || name.trim().equalsIgnoreCase(targetName.trim())) {
                return i;
            }
        }
        return -1;
    }

    private String getUniqueFileName(String desiredName, Set<String> usedSet) {
        if (!usedSet.contains(desiredName)) {
            usedSet.add(desiredName);
            return desiredName;
        }
        int dotIdx = desiredName.lastIndexOf('.');
        String base = dotIdx != -1 ? desiredName.substring(0, dotIdx) : desiredName;
        String ext = dotIdx != -1 ? desiredName.substring(dotIdx) : "";

        int counter = 1;
        while (usedSet.contains(base + "_(" + counter + ")" + ext)) {
            counter++;
        }
        String unique = base + "_(" + counter + ")" + ext;
        usedSet.add(unique);
        return unique;
    }

    private String sanitizeFileName(String name, String fallback) {
        String sanitized = name.replaceAll("[/\\\\?%*:|\"<>]", "_").trim();
        return !sanitized.isEmpty() ? sanitized : fallback;
    }

    private String sanitizeSheetTabName(String name) {
        String clean = name.replaceAll("[/\\\\?*\\[\\]:]", "_").trim();
        return clean.length() > 31 ? clean.substring(0, 31).trim() : clean;
    }

    private boolean isBinaryXlsFormat(byte[] bytes) {
        if (bytes == null || bytes.length < 8) return false;
        return (bytes[0] & 0xFF) == 0xD0 && (bytes[1] & 0xFF) == 0xCF;
    }

    private boolean isZipPackage(byte[] bytes) {
        if (bytes == null || bytes.length < 4) return false;
        return bytes[0] == 0x50 && bytes[1] == 0x4B;
    }

    /**
     * Sanitizes any Noto Sans Devanagari font fallbacks into Khmer OS Battambang.
     */
    private byte[] sanitizeDevanagariFonts(byte[] xlsxBytes) {
        try (ByteArrayInputStream bais = new ByteArrayInputStream(xlsxBytes);
             ZipInputStream zis = new ZipInputStream(bais);
             ByteArrayOutputStream baos = new ByteArrayOutputStream();
             ZipOutputStream zos = new ZipOutputStream(baos)) {

            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                String name = entry.getName();
                byte[] entryBytes = zis.readAllBytes();

                if (name.equals("xl/styles.xml") || name.equals("xl/sharedStrings.xml") || name.equals("xl/theme/theme1.xml")) {
                    String content = new String(entryBytes, StandardCharsets.UTF_8);
                    if (content.contains("Devanagari") || content.contains("Mangal")) {
                        content = content
                                .replaceAll("(?i)(val|typeface)=\"Noto Serif Devanagari\"", "$1=\"Khmer OS Muol Light\"")
                                .replaceAll("(?i)(val|typeface)=\"Noto Sans Devanagari(?:\\s+UI)?\"", "$1=\"Khmer OS Battambang\"")
                                .replaceAll("(?i)(val|typeface)=\"[^\"]*Devanagari[^\"]*\"", "$1=\"Khmer OS Battambang\"")
                                .replaceAll("(?i)(val|typeface)=\"Mangal\"", "$1=\"Khmer OS Battambang\"")
                                .replaceAll("(?i)>Noto Serif Devanagari<", ">Khmer OS Muol Light<")
                                .replaceAll("(?i)>Noto Sans Devanagari(?:\\s+UI)?<", ">Khmer OS Battambang<")
                                .replaceAll("(?i)>[^<]*Devanagari[^<]*<", ">Khmer OS Battambang<");
                        entryBytes = content.getBytes(StandardCharsets.UTF_8);
                    }
                }

                ZipEntry newEntry = new ZipEntry(name);
                zos.putNextEntry(newEntry);
                zos.write(entryBytes);
                zos.closeEntry();
            }

            zos.finish();
            return baos.toByteArray();
        } catch (Exception e) {
            log.warn("Font sanitization skipped due to error: {}", e.getMessage());
            return xlsxBytes;
        }
    }
}

