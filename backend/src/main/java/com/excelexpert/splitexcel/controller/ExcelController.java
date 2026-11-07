package com.excelexpert.splitexcel.controller;

import com.excelexpert.splitexcel.model.SplitConfigRequest;
import com.excelexpert.splitexcel.model.WorkbookMetaDto;
import com.excelexpert.splitexcel.service.ExcelSplitService;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.commons.io.output.ByteArrayOutputStream;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

@RestController
@RequestMapping("/api/excel")
public class ExcelController {

    private static final Logger log = LoggerFactory.getLogger(ExcelController.class);

    private final ExcelSplitService excelSplitService;
    private final ObjectMapper objectMapper;

    public ExcelController(ExcelSplitService excelSplitService, ObjectMapper objectMapper) {
        this.excelSplitService = excelSplitService;
        this.objectMapper = objectMapper;
    }

    /**
     * Inspects uploaded workbook (.xls or .xlsx) and returns sheet metadata and sample rows.
     */
    @PostMapping(value = "/inspect", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<WorkbookMetaDto> inspect(@RequestParam("file") MultipartFile file) throws IOException {
        String fileName = file.getOriginalFilename() != null ? file.getOriginalFilename() : "Workbook.xlsx";
        byte[] bytes = file.getBytes();
        WorkbookMetaDto meta = excelSplitService.inspectWorkbook(bytes, fileName);
        return ResponseEntity.ok(meta);
    }

    /**
     * Splits a single worksheet directly into an independent .xlsx file.
     */
    @PostMapping(value = "/split-single", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<byte[]> splitSingle(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetName") String sheetName,
            @RequestParam(value = "customName", required = false) String customName
    ) throws IOException {
        byte[] fileBytes = file.getBytes();
        byte[] sheetBytes = excelSplitService.splitSingleSheet(fileBytes, sheetName, customName);

        boolean isXls = (sheetBytes.length >= 8 && (sheetBytes[0] & 0xFF) == 0xD0 && (sheetBytes[1] & 0xFF) == 0xCF);
        String ext = isXls ? ".xls" : ".xlsx";
        String contentType = isXls ? "application/vnd.ms-excel" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        String outputName = customName != null && !customName.trim().isEmpty()
                ? customName.trim()
                : sheetName + ext;
        if (!outputName.toLowerCase().endsWith(ext)) {
            outputName += ext;
        }

        String encodedFilename = URLEncoder.encode(outputName, StandardCharsets.UTF_8).replaceAll("\\+", "%20");

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename*=UTF-8''" + encodedFilename)
                .contentType(MediaType.parseMediaType(contentType))
                .body(sheetBytes);
    }

    /**
     * Splits a single workbook into a ZIP containing its dedicated folder with all split sheets.
     */
    @PostMapping(value = "/split-zip", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<byte[]> splitZip(
            @RequestParam("file") MultipartFile file,
            @RequestParam(value = "config", required = false) String configJson
    ) throws IOException {
        SplitConfigRequest config = configJson != null && !configJson.trim().isEmpty()
                ? objectMapper.readValue(configJson, SplitConfigRequest.class)
                : new SplitConfigRequest(null, "", "-", "original", 1, "2", null, null, false);

        String originalName = file.getOriginalFilename() != null ? file.getOriginalFilename() : "Workbook.xlsx";
        byte[] zipBytes = excelSplitService.splitAndPackageZip(file.getBytes(), originalName, config);

        String baseName = originalName.replaceAll("\\.[^.]+$", "");
        String zipBaseName = config.folderName() != null && !config.folderName().trim().isEmpty()
                ? config.folderName().trim()
                : baseName;
        String zipName = zipBaseName + ".zip";
        String encodedZipName = URLEncoder.encode(zipName, StandardCharsets.UTF_8).replaceAll("\\+", "%20");

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename*=UTF-8''" + encodedZipName)
                .contentType(MediaType.parseMediaType("application/zip"))
                .body(zipBytes);
    }

    /**
     * Splits multiple master workbooks and bundles them into one master ZIP with individual master folders.
     */
    @PostMapping(value = "/split-batch-zip", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<byte[]> splitBatchZip(
            @RequestParam("files") List<MultipartFile> files,
            @RequestParam(value = "configs", required = false) String configsJson,
            @RequestParam(value = "globalConfig", required = false) String globalConfigJson
    ) throws IOException {
        Map<String, SplitConfigRequest> configMap = new HashMap<>();
        if (configsJson != null && !configsJson.trim().isEmpty()) {
            configMap = objectMapper.readValue(configsJson, new TypeReference<>() {});
        }

        SplitConfigRequest globalConfig = globalConfigJson != null && !globalConfigJson.trim().isEmpty()
                ? objectMapper.readValue(globalConfigJson, SplitConfigRequest.class)
                : new SplitConfigRequest(null, "", "-", "original", 1, "2", null, null, false);

        ByteArrayOutputStream masterZipBaos = new ByteArrayOutputStream();
        try (ZipOutputStream masterZos = new ZipOutputStream(masterZipBaos)) {
            Set<String> usedFolders = new HashSet<>();

            for (MultipartFile file : files) {
                String originalName = file.getOriginalFilename() != null ? file.getOriginalFilename() : "Workbook.xlsx";
                SplitConfigRequest fileConfig = configMap.getOrDefault(originalName, globalConfig);

                byte[] individualZipBytes = excelSplitService.splitAndPackageZip(file.getBytes(), originalName, fileConfig);

                // Transfer individual zip entries into master zip
                try (ZipInputStream zis = new ZipInputStream(new java.io.ByteArrayInputStream(individualZipBytes))) {
                    ZipEntry entry;
                    while ((entry = zis.getNextEntry()) != null) {
                        masterZos.putNextEntry(new ZipEntry(entry.getName()));
                        masterZos.write(zis.readAllBytes());
                        masterZos.closeEntry();
                    }
                }
            }
        }

        String masterZipName;
        if (globalConfig.folderName() != null && !globalConfig.folderName().trim().isEmpty()) {
            masterZipName = globalConfig.folderName().trim() + ".zip";
        } else if (files.size() == 1) {
            String firstBase = files.get(0).getOriginalFilename() != null
                    ? files.get(0).getOriginalFilename().replaceAll("\\.[^.]+$", "")
                    : "Workbook";
            masterZipName = firstBase + ".zip";
        } else if (!files.isEmpty()) {
            String firstBase = files.get(0).getOriginalFilename() != null
                    ? files.get(0).getOriginalFilename().replaceAll("\\.[^.]+$", "")
                    : "Workbook";
            masterZipName = firstBase + "_and_" + (files.size() - 1) + "_more.zip";
        } else {
            masterZipName = "Split_Excel_Sheets.zip";
        }
        String encoded = URLEncoder.encode(masterZipName, StandardCharsets.UTF_8).replaceAll("\\+", "%20");

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename*=UTF-8''" + encoded)
                .contentType(MediaType.parseMediaType("application/zip"))
                .body(masterZipBaos.toByteArray());
    }
}

