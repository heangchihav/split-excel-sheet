package com.excelexpert.splitexcel.model;

import java.util.List;

public record WorkbookMetaDto(
        String fileName,
        long fileSize,
        boolean isBinaryXls,
        List<SheetMetaDto> sheets
) {}

