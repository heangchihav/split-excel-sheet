package com.excelexpert.splitexcel.model;

import java.util.List;

public record SheetMetaDto(
        int index,
        String name,
        int rowCount,
        int colCount,
        List<String> sampleHeaders,
        List<List<String>> sampleRows
) {}

