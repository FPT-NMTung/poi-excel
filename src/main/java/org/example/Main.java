package org.example;

import io.vertx.core.json.JsonArray;
import io.vertx.core.json.JsonObject;
import model.*;
import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {
    public static void main(String[] args) throws Exception {
        // Get file template
        File templateFile = new File("template.xlsx");
        if (!templateFile.exists()) {
            throw new Exception("Template file not found");
        }

        long startTime;

        // Convert to POI
        XSSFWorkbook wb = new XSSFWorkbook(templateFile);

        // Get config setting
        ConfigSetting configSetting = getConfigSetting(wb);

        // Get JSON data
        String jsonStr = IOUtils.toString(new FileReader("./testData.json"));
        JsonArray sourceData = new JsonArray(jsonStr);

        // Process data
        System.out.println("Data size: " + sourceData.size() + " |" + (sourceData.size() >= 5000 ? " Warning: SLOWWW": ""));
        System.out.print("Start process data... ");
        startTime = System.currentTimeMillis();
        ChildTree childTree = processData(configSetting, sourceData);
        System.out.println((System.currentTimeMillis() - startTime) + "ms");


        // Generate file
        System.out.println("Start generate data... ");
        generateFile(wb, configSetting, childTree, sourceData);

        // Export file
        System.out.print("Export file... ");
        startTime = System.currentTimeMillis();
        FileOutputStream fOut = new FileOutputStream("./result.xlsx");
        wb.write(fOut);
        fOut.close();
        System.out.println((System.currentTimeMillis() - startTime) + "ms");

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

    private static ChildTree processData (ConfigSetting configSetting, JsonArray sourceData) {
        ChildTree resultData = new ChildTree();

        for (int index = 0; index < sourceData.size(); index++) {
            JsonObject itemData = sourceData.getJsonObject(index);

            processDataRecursive(index, resultData, itemData, 0, configSetting);
        }

        return resultData;
    }

    private static void processDataRecursive (int indexData, ChildTree childData, JsonObject itemData, int level, ConfigSetting configSetting) {
        // Condition break out recursive
        if (level >= configSetting.getTotalGroup()) {
            return;
        }

        if (childData.getData() == null) {
            childData.setLevel(level);
            childData.setData(new ArrayList<>());
            childData.setHashSetKey(new HashSet<>());
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
            ArrayList<ItemTree> listItemTree = childData.getData();
            HashSet<String> hashSetKey = childData.getHashSetKey();
            boolean hasKeyValue = hashSetKey.contains(keyString.toString());

            if (!hasKeyValue) {
                ItemTree newItemTree = new ItemTree();

                newItemTree.setValue(keyObject);
                newItemTree.setKey(keyString.toString());
                if (level + 1 < configSetting.getTotalGroup()) {
                    ChildTree childTree = new ChildTree();

                    processDataRecursive(indexData, childTree, itemData, level + 1, configSetting);

                    newItemTree.setChild(childTree);
                }

                hashSetKey.add(keyString.toString());
                listItemTree.add(newItemTree);
            } else {
                if (level + 1 < configSetting.getTotalGroup()) {
                    ItemTree itemTree = new ItemTree();

                    for (ItemTree item: listItemTree) {
                        if (Objects.equals(item.getKey(), keyString.toString())) {
                            itemTree = item;
                            break;
                        }
                    }

                    processDataRecursive(indexData, itemTree.getChild(), itemData, level + 1, configSetting);
                }
            }
        } else {
            ArrayList<ItemTree> listItemTree = childData.getData();

            ItemTree newItemTree = new ItemTree();
            newItemTree.setValue(itemData);
            newItemTree.setKey(String.valueOf(indexData));

            listItemTree.add(newItemTree);
        }
    }

    private static void generateFile (XSSFWorkbook sourceTemplate, ConfigSetting configSetting, ChildTree jsonArrData, JsonArray sourceData) throws Exception {
        // Get sheet data
        XSSFSheet sheet = sourceTemplate.getSheetAt(0);

        long startTime;

        // Move footer
        System.out.print("\tMove footer... ");
        startTime = System.currentTimeMillis();
        int heightTable = calculateTotalTableHeightRecursive(configSetting, jsonArrData, 0);
        sheet.shiftRows(new CellAddress(configSetting.getArrRange()[0].getEnd()).getRow() + 1, sheet.getLastRowNum(), heightTable, true, true);
        System.out.println((System.currentTimeMillis() - startTime) + "ms");

        // Init start row
        System.out.print("\tGenerate... ");
        startTime = System.currentTimeMillis();
        int startRow = new CellAddress(configSetting.getArrRange()[0].getEnd()).getRow() + 1;
        int totalAppend = generateFileFromTemplate(0, startRow, jsonArrData, sheet, configSetting);
        System.out.println((System.currentTimeMillis() - startTime) + "ms");

        // remove range template
        System.out.print("\tRemove template... ");
        startTime = System.currentTimeMillis();
        removeRangeTemplate(configSetting, sheet);
        System.out.println((System.currentTimeMillis() - startTime) + "ms");

        // remove config sheet
        System.out.print("\tRemove config sheet... ");
        startTime = System.currentTimeMillis();
        sourceTemplate.removeSheetAt(1);
        System.out.println((System.currentTimeMillis() - startTime) + "ms");

        // Fill data general
        if (configSetting.isHasGeneralData()) {
            System.out.print("\tFill data general... ");
            startTime = System.currentTimeMillis();
            fillDataGeneral(sourceTemplate, totalAppend, configSetting, sourceData.getJsonObject(0));
            System.out.println((System.currentTimeMillis() - startTime) + "ms");
        }

        // Merge cell
        if (configSetting.isMergeCell()) {
            startTime = System.currentTimeMillis();
            System.out.print("\tMerge cell... ");
            System.out.println((System.currentTimeMillis() - startTime) + "ms");
            mergeCell(sourceTemplate);
        }
    }

    private static int generateFileFromTemplate(int level, int startRow, ChildTree jsonArrData, XSSFSheet targetSheet, ConfigSetting configSetting) throws Exception {
        ArrayList<ItemTree> dataObject = jsonArrData.getData();

        int totalAppendRow = 0;

        for (ItemTree item: dataObject) {
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
                    fillData(range, item.getValue(), targetSheet, configSetting, "<#table.(.*?)>");

                    totalAppendRow += beginRowChildTemplate - beginRowParentTemplate;
//                    exportTempFile(targetSheet);
                }
            }

            // go to generate deep child (if level not highest - has child)
            if (level < configSetting.getTotalGroup() - 1) {
                ChildTree childObject = item.getChild();

                int numGenerateRowChild = generateFileFromTemplate(level + 1, startRow + totalAppendRow, childObject, targetSheet, configSetting);

                totalAppendRow += numGenerateRowChild;
//                exportTempFile(targetSheet);
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
                fillData(range, item.getValue(), targetSheet, configSetting, "<#table.(.*?)>");

                totalAppendRow += endRowChildTemplate - beginRowChildTemplate + 1;
//                exportTempFile(targetSheet);
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
                    fillData(range, item.getValue(), targetSheet, configSetting, "<#table.(.*?)>");

                    totalAppendRow += endRowParentTemplate - endRowChildTemplate;
//                    exportTempFile(targetSheet);
                }
            }
        }

        return totalAppendRow;
    }

    public static void exportTempFile(XSSFSheet targetSheet) throws Exception {
        FileOutputStream fOut = new FileOutputStream("./temp.xlsx");
        targetSheet.getWorkbook().write(fOut);
        fOut.close();
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

            fillData(range, firstRowData, sheet, configSetting, "<#general.(.*?)>");
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

            fillData(range, firstRowData, sheet, configSetting, "<#general.(.*?)>");
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
        sheet.shiftRows(rowNum + heightParent, lastRow, -heightParent, true, true);
    }

    private static void fillData (Range range, JsonObject data, XSSFSheet targetSheet, ConfigSetting configSetting, String regex) throws Exception {
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

                // fill data to cell
                if (valueCell != null && !valueCell.isEmpty()) {
                    Pattern pattern = Pattern.compile(regex);
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

                    String newValue = valueCell.replaceAll(matcher.group(0), value);

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

    private static int calculateTotalTableHeightRecursive(ConfigSetting configSetting, ChildTree data, int level) {
        int totalHeight = 0;

        // Get height only level
        int heightTemplateLevel = configSetting.getArrRange()[level].getHeightRange();
        if (level + 1 < configSetting.getTotalGroup()) {
            heightTemplateLevel -= configSetting.getArrRange()[level + 1].getHeightRange();
        }

        int sizeObject = data.getData().size();
        ArrayList<ItemTree> dataObject = data.getData();

        if (level + 1 < configSetting.getTotalGroup()) {
            for (ItemTree item: dataObject) {
                ChildTree childData = item.getChild();

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

                String valueCell = "";
                try {
                    valueCell = cell.getStringCellValue();
                } catch (Exception e) {
                    continue;
                }

                Pattern pattern = Pattern.compile("<#merge.(.*?)>");
                Matcher matcher = pattern.matcher(valueCell);

                String key;
                if (matcher.find()) {
                    key = matcher.group(1);
                } else {
                    continue;
                }

                String newValue = valueCell.replaceAll(matcher.group(0), "");

                cell.setCellValue(newValue);

                boolean hasKey = mergeCellLists.containsKey(matcher.group(1));
                if (!hasKey) {
                    MergeCellList mergeCellList = new MergeCellList(matcher.group(1));
                    mergeCellList.addCell(new CellAddress(cell));
                    mergeCellLists.put(matcher.group(1), mergeCellList);
                } else {
                    MergeCellList mergeCellList = mergeCellLists.get(matcher.group(1));
                    mergeCellList.addCell(new CellAddress(cell));
                    mergeCellLists.put(matcher.group(1), mergeCellList);
                }
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