package com.excelexpert.splitexcel.model;

import java.util.List;
import java.util.Map;

public record SplitConfigRequest(
        String folderName,
        String prefix,
        String separator,
        String namingMode,
        Integer startNumber,
        String zeroPadding,
        List<String> selectedSheets,
        Map<String, String> customSheetNames,
        Boolean firstSheetIsInfo,
        String infoFileName
) {
    public SplitConfigRequest(
            String folderName,
            String prefix,
            String separator,
            String namingMode,
            Integer startNumber,
            String zeroPadding,
            List<String> selectedSheets,
            Map<String, String> customSheetNames,
            Boolean firstSheetIsInfo
    ) {
        this(folderName, prefix, separator, namingMode, startNumber, zeroPadding, selectedSheets, customSheetNames, firstSheetIsInfo, null);
    }

    public String resolvePrefix() {
        return prefix != null ? prefix.trim() : "";
    }

    public String resolveSeparator() {
        return separator != null ? separator : "-";
    }

    public String resolveNamingMode() {
        return namingMode != null ? namingMode : "original";
    }

    public int resolveStartNumber() {
        return startNumber != null ? startNumber : 1;
    }

    public String resolveZeroPadding() {
        return zeroPadding != null ? zeroPadding : "2";
    }

    public boolean resolveFirstSheetIsInfo() {
        return firstSheetIsInfo != null && firstSheetIsInfo;
    }

    public String resolveInfoFileName() {
        return infoFileName != null ? infoFileName.trim() : "";
    }
}

