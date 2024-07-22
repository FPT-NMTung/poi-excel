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

import static converter.ExcelToPDFConverter.convertExcelToPDF;

public class Main {
    public static void main(String[] args) throws Exception {
        // Get file template
        File templateFile = new File("template-non-multi.xlsx");
//        File templateFile = new File("GD_07.xlsx");

        if (!templateFile.exists()) {
            throw new Exception("Template file not found");
        }

        long startTime;
        long beginTime = System.currentTimeMillis();

        // Convert to POI
        XSSFWorkbook wb = new XSSFWorkbook(templateFile);

        // Get config setting
        ConfigSetting configSetting = getConfigSetting(wb);

        // Loop sheet config
        for (int indexSheet = 0; indexSheet < configSetting.getSheets().size(); indexSheet++) {
            // get target sheet
            XSSFSheet targetSheet = wb.getSheetAt(indexSheet);
            SheetConfig sheetConfig = configSetting.getSheets().get(indexSheet);

            // Get JSON data
            String jsonStr = IOUtils.toString(new FileReader("./testData.json"));
            JsonArray sourceData = new JsonArray(jsonStr);

            // Process data
            System.out.println("Data size: " + sourceData.size() + " |" + (sourceData.size() >= 5000 ? " Warning: SLOWWW": ""));
            System.out.print("Start process data... ");
            startTime = System.currentTimeMillis();
            ChildTree childTree = processData(sheetConfig, configSetting, sourceData);
            System.out.println((System.currentTimeMillis() - startTime) + "ms");

            // Generate file
            System.out.println("Start generate data... ");
            generateFile(wb, targetSheet, configSetting, sheetConfig, childTree, sourceData);
        }

        // remove config sheet
        System.out.print("\tRemove config sheet... ");
        startTime = System.currentTimeMillis();
        wb.removeSheetAt(wb.getSheetIndex("config"));
        System.out.println((System.currentTimeMillis() - startTime) + "ms");

        // Export file
        System.out.print("Export file... ");
        startTime = System.currentTimeMillis();
        FileOutputStream fOut = new FileOutputStream("./result.xlsx");
        wb.write(fOut);
        fOut.close();
        System.out.println((System.currentTimeMillis() - startTime) + "ms");

        String excelFilePath = "./result.xlsx";
        String pdfFilePath = "./result.pdf";
        convertExcelToPDF(excelFilePath, pdfFilePath);

        System.out.println("Export done!");
        System.out.println("Total time: " + (System.currentTimeMillis() - beginTime) + "!");
    }

    private static ConfigSetting getConfigSetting (XSSFWorkbook wb) throws Exception {
        // Get sheet config
        XSSFSheet sheet = wb.getSheet("config");

        if (sheet == null) {
            throw new Exception("No config sheet found");
        }

        ConfigSetting configSetting = new ConfigSetting();

        // flag check config
        boolean isHasGeneralData    = false;
        boolean isHasMergeCell      = false;
        boolean isMultipleSheet     = false;

        for (int countRow = 0; countRow <= sheet.getLastRowNum(); countRow++) {
            XSSFRow row             = sheet.getRow(countRow);

            if (row == null) {
                break;
            }

            XSSFCell cell           = row.getCell(0);

            String nameConfigCol    = cell.getStringCellValue();
            XSSFCell cellValue      = row.getCell(1);
            int valueConfigCol      = (int) cellValue.getNumericCellValue();

            // skip row empty first column
            if (nameConfigCol.isBlank()) {
                continue;
            }

            // get setting global config
            if (!isHasGeneralData) {
                if (nameConfigCol.trim().equals("isHasGeneralData") && valueConfigCol == 1) {
                    isHasGeneralData = true;
                    configSetting.setHasGeneralData(true);
                }
            }
            if (!isHasMergeCell) {
                if (nameConfigCol.trim().equals("isMergeCell") && valueConfigCol == 1) {
                    isHasMergeCell = true;
                    configSetting.setMergeCell(true);
                }
            }
            if (!isMultipleSheet) {
                if (nameConfigCol.trim().equals("isMultipleSheet") && valueConfigCol == 1) {
                    isMultipleSheet = true;
                    configSetting.setMultipleSheet(true);
                }
            }
        }

        // get config table with non-multiple sheet
        if (!isMultipleSheet) {
            // init flag count non-multiple sheet
            int count                           = 0;
            ArrayList<SheetConfig> sheetConfigs = new ArrayList<>();
            SheetConfig sheetConfig             = new SheetConfig();

            sheetConfig.setIndex(0);

            for (int countRow = 0; countRow <= sheet.getLastRowNum(); countRow++) {
                XSSFRow row             = sheet.getRow(countRow);

                // skip row null
                if (row == null) {
                    continue;
                }

                XSSFCell cell           = row.getCell(0);

                // skip cell null
                if (cell == null) {
                    continue;
                }

                String nameConfigCol   = cell.getStringCellValue();

                // skip row empty first column
                if (nameConfigCol.isBlank()) {
                    continue;
                }

                // get config table with non-multiple sheet
                if (nameConfigCol.contains("range_")) {
                    String begin = row.getCell(1).getStringCellValue();
                    String end = row.getCell(2).getStringCellValue();
                    String columnData = row.getCell(3).getStringCellValue();

                    String[] columns = columnData.split(",");

                    Range range = new Range(begin, end, columns);

                    sheetConfig.getArrRange().add(range);
                }
            }

            sheetConfigs.add(sheetConfig);
            configSetting.setSheets(sheetConfigs);
        } else {
            // init flag count multiple sheet
            int currentSheetIndex               = -1;
            ArrayList<SheetConfig> sheetConfigs = new ArrayList<>();
            SheetConfig sheetConfig             = null;
            ArrayList<Range> ranges             = new ArrayList<>();

            for (int countRow = 0; countRow <= sheet.getLastRowNum(); countRow++) {
                XSSFRow row = sheet.getRow(countRow);

                // skip row null
                if (row == null) {
                    continue;
                }

                XSSFCell cell = row.getCell(0);
                String nameConfigCol    = cell.getStringCellValue();

                // skip row not config sheet and table
                if (!nameConfigCol.contains("range_") && !nameConfigCol.contains("sheet_")) {
                    continue;
                }

                // row config sheet
                if (nameConfigCol.contains("sheet_")) {
                    if (currentSheetIndex != -1) {
                        sheetConfigs.add(sheetConfig);
                    }
                    currentSheetIndex++;

                    sheetConfig = new SheetConfig();
                    ranges = new ArrayList<>();
                    sheetConfig.setArrRange(ranges);
                    sheetConfig.setIndex(currentSheetIndex);
                }

                // row config table
                if (nameConfigCol.contains("range_") && sheetConfig != null) {
                    String begin = row.getCell(1).getStringCellValue();
                    String end = row.getCell(2).getStringCellValue();
                    String columnData = row.getCell(3).getStringCellValue();

                    String[] columns = columnData.split(",");

                    Range range = new Range(begin, end, columns);

                    ranges.add(range);
                }
            }

            sheetConfigs.add(sheetConfig);
            configSetting.setSheets(sheetConfigs);
        }



//        boolean isHasGeneralData = (int) sheet.getRow(1).getCell(1).getNumericCellValue() == 1;
//        boolean isMergeCell = (int) sheet.getRow(2).getCell(1).getNumericCellValue() == 1;
//
//        ConfigSetting configSetting = new ConfigSetting(totalGroup, isHasGeneralData, isMergeCell);
//        Range[] ranges = configSetting.getArrRange();
//
//        // Get array object range
//        int countRow = 0;
//        int count = 0;
//
//        while (count < totalGroup || countRow < sheet.getLastRowNum()) {
//            XSSFRow row = sheet.getRow(countRow);
//            countRow += 1;
//
//            // Check null, exist content in row and cell
//            if (row == null) {
//                continue;
//            }
//
//            XSSFCell cell = row.getCell(0);
//            if (cell == null) {
//                continue;
//            }
//
//            String content = cell.getStringCellValue();
//
//            // Check and get information config
//            if (content.contains("range_" + count)) {
//                String begin = row.getCell(1).getStringCellValue();
//                String end = row.getCell(2).getStringCellValue();
//                String columnData = row.getCell(3).getStringCellValue();
//
//                String[] columns = columnData.split(",");
//
//                Range range = new Range(begin, end, columns);
//
//                ranges[count] = range;
//                count += 1;
//            }
//        }
//
//        configSetting.setArrRange(ranges);

        return configSetting;
    }

    private static ChildTree processData (SheetConfig sheetConfig, ConfigSetting configSetting, JsonArray sourceData) {
        ChildTree resultData = new ChildTree();

        int startIndex = 0;
        if (configSetting.isHasGeneralData()) {
            startIndex = 1;
        }

        for (int index = startIndex; index < sourceData.size(); index++) {
            JsonObject itemData = sourceData.getJsonObject(index);

            processDataRecursive(index, resultData, itemData, 0, configSetting, sheetConfig);
        }

        return resultData;
    }

    private static void processDataRecursive (int indexData, ChildTree childData, JsonObject itemData, int level, ConfigSetting configSetting, SheetConfig sheetConfig) {
        // Condition break out recursive
        if (level >= sheetConfig.getTotalGroup()) {
            return;
        }

        if (childData.getData() == null) {
            childData.setLevel(level);
            childData.setData(new ArrayList<>());
            childData.setHashSetKey(new HashSet<>());
        }

        // Get data follow config file
        String[] columnData = sheetConfig.getArrRange().get(level).getColumnData();
        boolean isColumnDataIsEmpty = sheetConfig.getArrRange().get(level).isColumnDataIsEmpty();

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
                if (level + 1 < sheetConfig.getTotalGroup()) {
                    ChildTree childTree = new ChildTree();

                    processDataRecursive(indexData, childTree, itemData, level + 1, configSetting, sheetConfig);

                    newItemTree.setChild(childTree);
                }

                hashSetKey.add(keyString.toString());
                listItemTree.add(newItemTree);
            } else {
                if (level + 1 < sheetConfig.getTotalGroup()) {
                    ItemTree itemTree = new ItemTree();

                    for (ItemTree item: listItemTree) {
                        if (Objects.equals(item.getKey(), keyString.toString())) {
                            itemTree = item;
                            break;
                        }
                    }

                    processDataRecursive(indexData, itemTree.getChild(), itemData, level + 1, configSetting, sheetConfig);
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

    private static void generateFile (XSSFWorkbook sourceTemplate, XSSFSheet targetSheet, ConfigSetting configSetting, SheetConfig sheetConfig, ChildTree jsonArrData, JsonArray sourceData) throws Exception {
        long startTime;

        sourceTemplate.setForceFormulaRecalculation(true);
        XSSFFormulaEvaluator.evaluateAllFormulaCells(sourceTemplate);

        // Move footer
        System.out.print("\tMove footer... ");
        startTime = System.currentTimeMillis();
        int heightTable = calculateTotalTableHeightRecursive(sheetConfig, jsonArrData, 0);
        if (heightTable > 0) {
            targetSheet.shiftRows(new CellAddress(sheetConfig.getArrRange().getFirst().getEnd()).getRow() + 1, targetSheet.getLastRowNum(), heightTable, true, true);
        }
        System.out.println((System.currentTimeMillis() - startTime) + "ms");

        // Init start row
        int startRow = new CellAddress(sheetConfig.getArrRange().getFirst().getEnd()).getRow() + 1;
        int totalAppend = 0;
        if (heightTable > 0) {
            System.out.print("\tGenerate... ");
            startTime = System.currentTimeMillis();
            totalAppend = generateFileFromTemplate(0, startRow, jsonArrData, targetSheet, configSetting, sheetConfig);
            System.out.println((System.currentTimeMillis() - startTime) + "ms");
        }

        // remove range template
        System.out.print("\tRemove template... ");
        startTime = System.currentTimeMillis();
        removeRangeTemplate(sheetConfig, targetSheet);
        System.out.println((System.currentTimeMillis() - startTime) + "ms");

        // Fill data general
        if (configSetting.isHasGeneralData()) {
            System.out.print("\tFill data general... ");
            startTime = System.currentTimeMillis();
            fillDataGeneral(targetSheet, totalAppend, configSetting, sheetConfig, sourceData.getJsonObject(0));
            System.out.println((System.currentTimeMillis() - startTime) + "ms");
        }

        // Merge cell
        if (configSetting.isMergeCell() && totalAppend > 0) {
            startTime = System.currentTimeMillis();
            System.out.print("\tMerge cell... ");
            System.out.println((System.currentTimeMillis() - startTime) + "ms");
            mergeCell(targetSheet);
        }

        XSSFFormulaEvaluator.evaluateAllFormulaCells(sourceTemplate);
    }

    private static int generateFileFromTemplate(int level, int startRow, ChildTree jsonArrData, XSSFSheet targetSheet, ConfigSetting configSetting, SheetConfig sheetConfig) throws Exception {
        ArrayList<ItemTree> dataObject = jsonArrData.getData();

        int totalAppendRow = 0;

        for (ItemTree item: dataObject) {
            // generate header row between parent and child (if level not highest)
            if (level < sheetConfig.getTotalGroup() - 1) {
                int beginRowParentTemplate = new CellAddress(sheetConfig.getArrRange().get(level).getBegin()).getRow();
                int beginRowChildTemplate = new CellAddress(sheetConfig.getArrRange().get(level + 1).getBegin()).getRow();

                // Only generate if it has content between
                if (beginRowChildTemplate - beginRowParentTemplate > 0) {
                    targetSheet.copyRows(beginRowParentTemplate, beginRowChildTemplate - 1, startRow + totalAppendRow, new CellCopyPolicy());

                    // Fill data to row
                    CellAddress beginRange = new CellAddress(startRow + totalAppendRow, 0);
                    CellAddress endRange = new CellAddress(beginRange.getRow() + beginRowChildTemplate - beginRowParentTemplate - 1, new CellAddress(sheetConfig.getArrRange().getFirst().getEnd()).getColumn());
                    Range range = new Range(beginRange.toString(), endRange.toString());
                    fillData(range, item.getValue(), targetSheet, configSetting, "<#table.(.*?)>");

                    totalAppendRow += beginRowChildTemplate - beginRowParentTemplate;
//                    exportTempFile(targetSheet);
                }
            }

            // go to generate deep child (if level not highest - has child)
            if (level < sheetConfig.getTotalGroup() - 1) {
                ChildTree childObject = item.getChild();

                int numGenerateRowChild = generateFileFromTemplate(level + 1, startRow + totalAppendRow, childObject, targetSheet, configSetting, sheetConfig);

                totalAppendRow += numGenerateRowChild;
//                exportTempFile(targetSheet);
            }

            // generate row chill
            if (level == sheetConfig.getTotalGroup() - 1) {
                int beginRowChildTemplate = new CellAddress(sheetConfig.getArrRange().get(level).getBegin()).getRow();
                int endRowChildTemplate = new CellAddress(sheetConfig.getArrRange().get(level).getEnd()).getRow();

                targetSheet.copyRows(beginRowChildTemplate, endRowChildTemplate, startRow + totalAppendRow, new CellCopyPolicy());

                // fill data to row
                CellAddress beginRange = new CellAddress(startRow + totalAppendRow, 0);
                CellAddress endRange = new CellAddress(beginRange.getRow() + sheetConfig.getArrRange().get(level).getHeightRange() - 1, new CellAddress(sheetConfig.getArrRange().getFirst().getEnd()).getColumn());
                Range range = new Range(beginRange.toString(), endRange.toString());
                fillData(range, item.getValue(), targetSheet, configSetting, "<#table.(.*?)>");

                totalAppendRow += endRowChildTemplate - beginRowChildTemplate + 1;
//                exportTempFile(targetSheet);
            }

            // generate footer row between parent and child (if level not highest)
            if (level < sheetConfig.getTotalGroup() - 1) {
                int endRowParentTemplate = new CellAddress(sheetConfig.getArrRange().get(level).getEnd()).getRow();
                int endRowChildTemplate = new CellAddress(sheetConfig.getArrRange().get(level + 1).getEnd()).getRow();

                // Only generate if it has content between
                if (endRowParentTemplate - endRowChildTemplate > 0) {
                    targetSheet.copyRows(endRowChildTemplate + 1, endRowParentTemplate, startRow + totalAppendRow, new CellCopyPolicy());

                    // Fill data to row
                    CellAddress beginRange = new CellAddress(startRow + totalAppendRow, 0);
                    CellAddress endRange = new CellAddress(beginRange.getRow() + endRowParentTemplate - endRowChildTemplate - 1, new CellAddress(sheetConfig.getArrRange().getFirst().getEnd()).getColumn());
                    Range range = new Range(beginRange.toString(), endRange.toString());
                    fillData(range, item.getValue(), targetSheet, configSetting, "<#table.(.*?)>");

                    totalAppendRow += endRowParentTemplate - endRowChildTemplate;
//                    exportTempFile(targetSheet);
                }
            }
        }

        return totalAppendRow;
    }

    private static void exportTempFile(XSSFSheet targetSheet) throws Exception {
        FileOutputStream fOut = new FileOutputStream("./temp.xlsx");
        targetSheet.getWorkbook().write(fOut);
        fOut.close();
    }

    private static void fillDataGeneral(XSSFSheet sheet, int totalAppend, ConfigSetting configSetting, SheetConfig sheetConfig, JsonObject firstRowData) throws Exception {
        int startRow = 0;
        int endRow = new CellAddress(sheetConfig.getArrRange().getFirst().getBegin()).getRow() - 1;

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

        startRow = new CellAddress(sheetConfig.getArrRange().getFirst().getBegin()).getRow() + totalAppend;
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

    private static void removeRangeTemplate (SheetConfig sheetConfig, XSSFSheet sheet) {
        int heightParent = sheetConfig.getArrRange().getFirst().getHeightRange();
        int rowNum = new CellAddress(sheetConfig.getArrRange().getFirst().getBegin()).getRow();

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

                CellType a = cell.getCellType();

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

    private static int calculateTotalTableHeightRecursive(SheetConfig sheetConfig, ChildTree data, int level) {
        int totalHeight = 0;

        // Get height only level
        int heightTemplateLevel = sheetConfig.getArrRange().get(level).getHeightRange();
        if (level + 1 < sheetConfig.getTotalGroup()) {
            heightTemplateLevel -= sheetConfig.getArrRange().get(level).getHeightRange();
        }

        if (data.getData() == null) {
            return 0;
        }

        int sizeObject = data.getData().size();
        ArrayList<ItemTree> dataObject = data.getData();

        if (level + 1 < sheetConfig.getTotalGroup()) {
            for (ItemTree item: dataObject) {
                ChildTree childData = item.getChild();

                int heightChildLevel = calculateTotalTableHeightRecursive(sheetConfig, childData, level + 1);
                totalHeight += heightChildLevel;
            }
        }

        return totalHeight + heightTemplateLevel * sizeObject;
    }

    private static void mergeCell(XSSFSheet sheet) {

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