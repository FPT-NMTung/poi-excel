package org.example;

import io.vertx.core.json.JsonArray;
import io.vertx.core.json.JsonObject;
import model.ConfigSetting;
import model.MergeCellList;
import model.Range;
import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {
    public static void main(String[] args) throws Exception {
        // Get file template
        File templateFile = new File("template.xlsx");
        if (!templateFile.exists()) {
            throw new Exception("Template file not found");
        }

        // Convert to POI
        XSSFWorkbook wb = new XSSFWorkbook(templateFile);

        // Get config setting
        ConfigSetting configSetting = getConfigSetting(wb);

        // Get JSON data
        String jsonStr = IOUtils.toString(new FileReader("./data.json"));
        JsonArray sourceData = new JsonArray(jsonStr);

        // Process data
        System.out.println("Data size: " + sourceData.size() + " |" + (sourceData.size() >= 5000 ? " Warning: SLOWWW": ""));
        System.out.println("Start process data...");
        JsonObject jsonArrData = processData(configSetting, sourceData);

        // Generate file
        System.out.println("Start generate data...");
        generateFile(wb, configSetting, jsonArrData, sourceData);

        // Export file
        System.out.println("Export file...");
        FileOutputStream fOut = new FileOutputStream("./result.xlsx");
        wb.write(fOut);
        fOut.close();

        System.out.println("Export done!");
    }

    private static ConfigSetting getConfigSetting (XSSFWorkbook wb) throws Exception {
        // Get sheet config
        XSSFSheet sheet = wb.getSheet("config");

        if (sheet == null) {
            throw new Exception("No config sheet found");
        }

        // Get total group
        int totalGroup = (int) sheet.getRow(0).getCell(1).getNumericCellValue();
        boolean isHasGeneralData = (int) sheet.getRow(1).getCell(1).getNumericCellValue() == 1;
        boolean isMergeCell = (int) sheet.getRow(2).getCell(1).getNumericCellValue() == 1;

        ConfigSetting configSetting = new ConfigSetting(totalGroup, isHasGeneralData, isMergeCell);
        Range[] ranges = configSetting.getArrRange();

        // Get array object range
        int countRow = 0;
        int count = 0;

        while (count < totalGroup || countRow < sheet.getLastRowNum()) {
            XSSFRow row = sheet.getRow(countRow);
            countRow += 1;

            // Check null, exist content in row and cell
            if (row == null) {
                continue;
            }

            XSSFCell cell = row.getCell(0);
            if (cell == null) {
                continue;
            }

            String content = cell.getStringCellValue();

            // Check and get information config
            if (content.contains("range_" + count)) {
                String begin = row.getCell(1).getStringCellValue();
                String end = row.getCell(2).getStringCellValue();
                String columnData = row.getCell(3).getStringCellValue();

                String[] columns = columnData.split(",");

                Range range = new Range(begin, end, columns);

                ranges[count] = range;
                count += 1;
            }
        }

        configSetting.setArrRange(ranges);

        return configSetting;
    }

    private static JsonObject processData (ConfigSetting configSetting, JsonArray sourceData) {
        JsonObject resultData = new JsonObject();

        for (int index = 0; index < sourceData.size(); index++) {
            JsonObject itemData = sourceData.getJsonObject(index);

            processDataRecursive(index, resultData, itemData, 0, configSetting);
        }

        return resultData;
    }

    private static void processDataRecursive (int indexData, JsonObject resultDataLevel, JsonObject itemData, int level, ConfigSetting configSetting) {
        // Condition break out recursive
        if (level >= configSetting.getTotalGroup()) {
            return;
        }

        if (resultDataLevel.encode().equals("{}")) {
            resultDataLevel.put("level", level);
            resultDataLevel.put("data", new JsonObject());
        }

        // Get data follow config file
        String[] columnData = configSetting.getArrRange()[level].getColumnData();
        boolean isColumnDataIsEmpty = configSetting.getArrRange()[level].isColumnDataIsEmpty();

        if (!isColumnDataIsEmpty) {
            StringBuilder keyString = new StringBuilder();
            JsonObject keyObject = new JsonObject();
            for (String columnDatum : columnData) {
                keyString.append(itemData.getValue(columnDatum).toString());
                keyObject.put(columnDatum, itemData.getValue(columnDatum).toString());
            }

            // Check exist key in result object
            JsonObject dataObject = resultDataLevel.getJsonObject("data");
            JsonObject findKeyString = dataObject.getJsonObject(keyString.toString());

            if (findKeyString == null) {
                dataObject.put(keyString.toString(), new JsonObject());

                JsonObject keyOb = dataObject.getJsonObject(keyString.toString());

                keyOb.put("value", keyObject);

                if (level + 1 < configSetting.getTotalGroup()) {
                    keyOb.put("child", new JsonObject());
                    processDataRecursive(indexData, keyOb.getJsonObject("child") ,itemData, level + 1, configSetting);
                }
            } else {
                if (level + 1 < configSetting.getTotalGroup()) {
                    processDataRecursive(indexData, findKeyString.getJsonObject("child") ,itemData, level + 1, configSetting);
                }
            }
        } else {
            JsonObject dataObject = resultDataLevel.getJsonObject("data");
            JsonObject valueObject = new JsonObject();
            valueObject.put("value", itemData);
            dataObject.put(String.valueOf(indexData), valueObject);
        }
    }

    private static void generateFile (XSSFWorkbook sourceTemplate, ConfigSetting configSetting, JsonObject jsonArrData, JsonArray sourceData) throws Exception {
        // Get sheet data
        XSSFSheet sheet = sourceTemplate.getSheetAt(0);

        // Move footer
        System.out.println("\tMove footer ...");
        int heightTable = calculateTotalTableHeightRecursive(configSetting, jsonArrData, 0);
        sheet.shiftRows(new CellAddress(configSetting.getArrRange()[0].getEnd()).getRow() + 1, sheet.getLastRowNum(), heightTable, true, true);

        // Init start row
        System.out.println("\tGenerate ...");
        int startRow = new CellAddress(configSetting.getArrRange()[0].getEnd()).getRow() + 1;
        int totalAppend = generateFileFromTemplate(0, startRow, jsonArrData, sheet, configSetting);

        // remove range template
        System.out.println("\tRemove template ...");
        removeRangeTemplate(configSetting, sheet);

        // remove config sheet
        System.out.println("\tRemove config sheet ...");
        sourceTemplate.removeSheetAt(1);

        // Fill data general
        if (configSetting.isHasGeneralData()) {
            System.out.println("\tFill data general ...");
            fillDataGeneral(sourceTemplate, totalAppend, configSetting, sourceData.getJsonObject(0));
        }

        // Merge cell
        if (configSetting.isMergeCell()) {
            System.out.println("\tMerge cell ...");
            mergeCell(sourceTemplate);
        }
    }

    private static int generateFileFromTemplate(int level, int startRow, JsonObject jsonArrData, XSSFSheet targetSheet, ConfigSetting configSetting) throws Exception {
        JsonObject dataObject = jsonArrData.getJsonObject("data");

        int totalAppendRow = 0;

        for (Map.Entry<String, Object> item: dataObject) {
            // generate header row between parent and child (if level not highest)
            if (level < configSetting.getTotalGroup() - 1) {
                int beginRowParentTemplate = new CellAddress(configSetting.getArrRange()[level].getBegin()).getRow();
                int beginRowChildTemplate = new CellAddress(configSetting.getArrRange()[level + 1].getBegin()).getRow();

                // Only generate if it has content between
                if (beginRowChildTemplate - beginRowParentTemplate > 0) {
                    targetSheet.copyRows(beginRowParentTemplate, beginRowChildTemplate - 1, startRow + totalAppendRow, new CellCopyPolicy());

                    // Fill data to row
                    CellAddress beginRange = new CellAddress(startRow + totalAppendRow, 0);
                    CellAddress endRange = new CellAddress(beginRange.getRow() + beginRowChildTemplate - beginRowParentTemplate - 1, new CellAddress(configSetting.getArrRange()[0].getEnd()).getColumn());
                    Range range = new Range(beginRange.toString(), endRange.toString());
                    fillData(range, ((JsonObject) item.getValue()).getJsonObject("value"), targetSheet, configSetting);

                    totalAppendRow += beginRowChildTemplate - beginRowParentTemplate;
                    exportTempFile(targetSheet);
                }
            }

            // go to generate deep child (if level not highest - has child)
            if (level < configSetting.getTotalGroup() - 1) {
                JsonObject childObject = ((JsonObject) item.getValue()).getJsonObject("child");

                int numGenerateRowChild = generateFileFromTemplate(level + 1, startRow + totalAppendRow, childObject, targetSheet, configSetting);

                totalAppendRow += numGenerateRowChild;
                exportTempFile(targetSheet);
            }

            // generate row chill
            if (level == configSetting.getTotalGroup() - 1) {
                int beginRowChildTemplate = new CellAddress(configSetting.getArrRange()[level].getBegin()).getRow();
                int endRowChildTemplate = new CellAddress(configSetting.getArrRange()[level].getEnd()).getRow();

                targetSheet.copyRows(beginRowChildTemplate, endRowChildTemplate, startRow + totalAppendRow, new CellCopyPolicy());

                // fill data to row
                CellAddress beginRange = new CellAddress(startRow + totalAppendRow, 0);
                CellAddress endRange = new CellAddress(beginRange.getRow() + configSetting.getArrRange()[level].getHeightRange() - 1, new CellAddress(configSetting.getArrRange()[0].getEnd()).getColumn());
                Range range = new Range(beginRange.toString(), endRange.toString());
                fillData(range, ((JsonObject) item.getValue()).getJsonObject("value"), targetSheet, configSetting);

                totalAppendRow += endRowChildTemplate - beginRowChildTemplate + 1;
                exportTempFile(targetSheet);
            }

            // generate footer row between parent and child (if level not highest)
            if (level < configSetting.getTotalGroup() - 1) {
                int endRowParentTemplate = new CellAddress(configSetting.getArrRange()[level].getEnd()).getRow();
                int endRowChildTemplate = new CellAddress(configSetting.getArrRange()[level + 1].getEnd()).getRow();

                // Only generate if it has content between
                if (endRowParentTemplate - endRowChildTemplate > 0) {
                    targetSheet.copyRows(endRowChildTemplate + 1, endRowParentTemplate, startRow + totalAppendRow, new CellCopyPolicy());

                    // Fill data to row
                    CellAddress beginRange = new CellAddress(startRow + totalAppendRow, 0);
                    CellAddress endRange = new CellAddress(beginRange.getRow() + endRowParentTemplate - endRowChildTemplate - 1, new CellAddress(configSetting.getArrRange()[0].getEnd()).getColumn());
                    Range range = new Range(beginRange.toString(), endRange.toString());
                    fillData(range, ((JsonObject) item.getValue()).getJsonObject("value"), targetSheet, configSetting);

                    totalAppendRow += endRowParentTemplate - endRowChildTemplate;
                    exportTempFile(targetSheet);
                }
            }
        }

        return totalAppendRow;
    }

    public static void exportTempFile(XSSFSheet targetSheet) throws Exception {
//        FileOutputStream fOut = new FileOutputStream("./temp.xlsx");
//        targetSheet.getWorkbook().write(fOut);
//        fOut.close();
    }

    private static void fillDataGeneral(XSSFWorkbook wb, int totalAppend, ConfigSetting configSetting, JsonObject firstRowData) throws Exception {
        XSSFSheet sheet = wb.getSheetAt(0);

        int startRow = 0;
        int endRow = new CellAddress(configSetting.getArrRange()[0].getBegin()).getRow() - 1;

        for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
            XSSFRow row = sheet.getRow(rowNum);
            if (row == null) {
                continue;
            }

            int lastCol = row.getLastCellNum();

            CellAddress cellBegin = new CellAddress(rowNum, 0);
            CellAddress cellEnd = new CellAddress(rowNum, lastCol);

            Range range = new Range(cellBegin.toString(), cellEnd.toString());

            fillData(range, firstRowData, sheet, configSetting);
        }

        startRow = new CellAddress(configSetting.getArrRange()[0].getBegin()).getRow() + totalAppend;
        endRow = sheet.getLastRowNum();

        for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
            XSSFRow row = sheet.getRow(rowNum);
            if (row == null) {
                continue;
            }

            int lastCol = row.getLastCellNum();

            CellAddress cellBegin = new CellAddress(rowNum, 0);
            CellAddress cellEnd = new CellAddress(rowNum, lastCol);

            Range range = new Range(cellBegin.toString(), cellEnd.toString());

            fillData(range, firstRowData, sheet, configSetting);
        }
    }

    private static void removeRangeTemplate (ConfigSetting configSetting, XSSFSheet sheet) {
        int heightParent = configSetting.getArrRange()[0].getHeightRange();
        int rowNum = new CellAddress(configSetting.getArrRange()[0].getBegin()).getRow();

        for (int index = 0; index < heightParent; index++) {
            XSSFRow row = sheet.getRow(index + rowNum);
            sheet.removeRow(row);
        }

        int lastRow = sheet.getLastRowNum();
        sheet.shiftRows(rowNum + heightParent, lastRow, -heightParent);
    }

    private static void fillData (Range range, JsonObject data, XSSFSheet targetSheet, ConfigSetting configSetting) throws Exception {
        int rowStart = new CellAddress(range.getBegin()).getRow();
        int rowEnd = new CellAddress(range.getEnd()).getRow();

        int colStart = new CellAddress(range.getBegin()).getColumn();
        int colEnd = new CellAddress(range.getEnd()).getColumn();

        for (int row = rowStart; row <= rowEnd; row++) {
            XSSFRow rowData = targetSheet.getRow(row);
            if (rowData == null) {
                continue;
            }

            for (int col = colStart; col <= colEnd; col++) {
                XSSFCell cell = rowData.getCell(col);
                if (cell == null) {
                    continue;
                }

                String valueCell = "";
                try {
                    valueCell = cell.getStringCellValue();
                } catch (Exception e) {
                    continue;
                }

                if (valueCell != null && !valueCell.isEmpty()) {
                    Pattern pattern = Pattern.compile("<#(.*?)>");
                    Matcher matcher = pattern.matcher(valueCell);

                    String key;
                    if (matcher.find()) {
                        key = matcher.group(1);
                    } else {
                        continue;
                    }

                    String value = data.getString(key);
                    if (value == null) {
                        continue;
                    }

                    String newValue = valueCell.replaceAll("<#(.*?)>", value);

                    XSSFComment comment = cell.getCellComment();
                    if (comment == null) {
                        continue;
                    }

                    String commentValue = String.valueOf(comment.getString());
                    if (!commentValue.isEmpty()) {
                        System.out.println(commentValue);
                    }

                    // Check format cell
                    if (cell.getCellStyle().getDataFormat() == 3) {
                        try {
                            double doubleValue = Double.parseDouble(newValue);
                            cell.setCellValue(doubleValue);
                        } catch (Exception e) {
                            cell.setCellValue(newValue);
                        }
                    } else {
                        cell.setCellValue(newValue);
                    }
                }
            }
        }
    }

    private static int calculateTotalTableHeightRecursive(ConfigSetting configSetting, JsonObject data, int level) {
        int totalHeight = 0;

        // Get height only level
        int heightTemplateLevel = configSetting.getArrRange()[level].getHeightRange();
        if (level + 1 < configSetting.getTotalGroup()) {
            heightTemplateLevel -= configSetting.getArrRange()[level + 1].getHeightRange();
        }

        int sizeObject = data.getJsonObject("data").size();
        JsonObject dataObject = data.getJsonObject("data");

        if (level + 1 < configSetting.getTotalGroup()) {
            for (Map.Entry<String, Object> item: dataObject) {
                JsonObject childData = ((JsonObject) item.getValue()).getJsonObject("child");

                int heightChildLevel = calculateTotalTableHeightRecursive(configSetting, childData, level + 1);
                totalHeight += heightChildLevel;
            }
        }

        return totalHeight + heightTemplateLevel * sizeObject;
    }

    private static void mergeCell(XSSFWorkbook wb) {
        XSSFSheet sheet = wb.getSheetAt(0);

        HashMap<String, MergeCellList> mergeCellLists = new HashMap<>();

        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            XSSFRow row = sheet.getRow(i);
            if (row == null) {
                continue;
            }

            for (int j = 0; j < row.getLastCellNum(); j++) {
                XSSFCell cell = row.getCell(j);
                if (cell == null) {
                    continue;
                }

                XSSFComment comment = cell.getCellComment();
                if (comment == null) {
                    continue;
                }

                String commentValue = String.valueOf(comment.getString());
                boolean hasKey = mergeCellLists.containsKey(commentValue);
                if (!hasKey) {
                    MergeCellList mergeCellList = new MergeCellList(commentValue);
                    mergeCellList.addCell(new CellAddress(cell));
                    mergeCellLists.put(commentValue, mergeCellList);
                } else {
                    MergeCellList mergeCellList = mergeCellLists.get(commentValue);
                    mergeCellList.addCell(new CellAddress(cell));
                    mergeCellLists.put(commentValue, mergeCellList);
                }

                String valueCell = cell.getStringCellValue();
                if (valueCell != null && commentValue.contains("(empty)")) {
                    cell.setCellValue("");
                }
//                cell.removeCellComment();
            }
        }

        for (Map.Entry<String, MergeCellList> item: mergeCellLists.entrySet()) {
            MergeCellList mergeCellList = item.getValue();

            if (mergeCellList.getCells().size() >= 2) {
                CellRangeAddress cellAddresses = mergeCellList.getCellRangeAddress();

                sheet.addMergedRegion(cellAddresses);
            }
        }

    }
}