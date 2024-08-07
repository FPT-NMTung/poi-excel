package org.example;

import io.vertx.core.json.JsonArray;
import io.vertx.core.json.JsonObject;
import model.*;
import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellType;
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
        File templateFile = new File("template-equivalent-tables.xlsx");
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
            String jsonStr = IOUtils.toString(new FileReader("./testDataSimple.json"));
            JsonArray sourceData = new JsonArray(jsonStr);

            // Process data
            System.out.println("Data size: " + sourceData.size() + " |" + (sourceData.size() >= 5000 ? " Warning: SLOWWW": ""));
            System.out.print("Start process data... ");
            startTime = System.currentTimeMillis();
            LevelDataTable rootLevelDataTable = processData(sheetConfig, configSetting, sourceData);
            System.out.println((System.currentTimeMillis() - startTime) + "ms");

            // Generate file
            System.out.println("Start generate data... ");
            generateFile(wb, targetSheet, configSetting, sheetConfig, rootLevelDataTable, sourceData);
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

//        String excelFilePath = "./result.xlsx";
//        String pdfFilePath = "./result.pdf";
//        convertExcelToPDF(excelFilePath, pdfFilePath);

        System.out.println("Export done!");
        System.out.println("Total time: " + (System.currentTimeMillis() - beginTime) + "ms!");
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
                    processRangeConfig(row, sheetConfig.getArrRange());
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
                    processRangeConfig(row, sheetConfig.getArrRange());
                }
            }

            sheetConfigs.add(sheetConfig);
            configSetting.setSheets(sheetConfigs);
        }

        return configSetting;
    }

    private static void processRangeConfig(XSSFRow row, ArrayList<Range> ranges) {
        String nameRow = row.getCell(0).getStringCellValue();
        String begin = row.getCell(1).getStringCellValue();
        String end = row.getCell(2).getStringCellValue();
        String columnData = row.getCell(3).getStringCellValue();

        XSSFCell cell = row.getCell(4);

        String indexTableExcel = null;
        String columnIndexTableExcel = null;
        if (cell != null) {
            String valueIndexTableExcel = row.getCell(4).getStringCellValue();

            if (!valueIndexTableExcel.isBlank()) {
                indexTableExcel = valueIndexTableExcel.split("\\|")[1];
                columnIndexTableExcel = valueIndexTableExcel.split("\\|")[0];
            }
        }

        String[] columns = columnData.split(",");

        // create range row
        Range range = new Range(begin, end, columns, indexTableExcel, columnIndexTableExcel);

        // get level row
        int targetLevel = Integer.parseInt(nameRow.replace("range_", ""));

        processRangeConfigRecursive(range, ranges, 0, targetLevel);
    }

    private static void processRangeConfigRecursive(Range range, ArrayList<Range> currentRanges, int currentLevel, int targetLevel) {
        if (currentLevel < targetLevel) {
            ArrayList<Range> childRanges = currentRanges.getLast().getChildRange();

            // Check null child range
            if (childRanges == null) {
                childRanges = new ArrayList<>();
                currentRanges.getLast().setChildRange(childRanges);
            }

            processRangeConfigRecursive(range, childRanges, currentLevel + 1, targetLevel);
            return;
        }

        currentRanges.add(range);
    }

    private static LevelDataTable processData (SheetConfig sheetConfig, ConfigSetting configSetting, JsonArray sourceData) {
        LevelDataTable rootLevelDataTable = new LevelDataTable(0);

        int startIndex = 0;
        if (configSetting.isHasGeneralData()) {
            startIndex = 1;
        }

        for (int index = startIndex; index < sourceData.size(); index++) {
            JsonObject itemData = sourceData.getJsonObject(index);

            processDataRecursive(rootLevelDataTable, itemData, 0, sheetConfig, sheetConfig.getArrRange());
        }

        return rootLevelDataTable;
    }

    private static void processDataRecursive (LevelDataTable levelDataTable, JsonObject itemData, int level, SheetConfig sheetConfig, ArrayList<Range> currentRangeConfig) {
        // check range has index table excel
        if (currentRangeConfig.getFirst().getIndexTableExcel() != null && !currentRangeConfig.getFirst().getIndexTableExcel().isEmpty()) {
            String dataIndexTableExcel = itemData.getString(currentRangeConfig.getFirst().getColumnIndexTableExcel());

            // Find dataTable with same indexRowTable
            DataTable selectedDataTable = null;
            for (int index = 0; index < levelDataTable.getDataTables().size(); index++) {
                if (levelDataTable.getDataTables().get(index).getIndexTableExcel().equals(dataIndexTableExcel)) {
                    selectedDataTable = levelDataTable.getDataTables().get(index);
                    break;
                }
            }

            // if not exist dataTable => create new
            if (selectedDataTable == null) {
                selectedDataTable = new DataTable();
                selectedDataTable.setIndexTableExcel(dataIndexTableExcel);
                levelDataTable.getDataTables().add(selectedDataTable);
            }

            // Find current range config
            Range selectedRangeConfig = null;
            for (int index = 0; index < currentRangeConfig.size(); index++) {
                if (currentRangeConfig.get(index).getIndexTableExcel().equals(dataIndexTableExcel)) {
                    selectedRangeConfig = currentRangeConfig.get(index);
                }
            }

            assert selectedRangeConfig != null;
            String[] columnData = selectedRangeConfig.getColumnData();

            // Check and process with no group column => leaf data
            if (columnData == null || columnData.length == 0 || (columnData.length == 1 && columnData[0].isEmpty())) {
                RowData newRowData = new RowData(itemData, null, level);

                selectedDataTable.getRowData().add(newRowData);
            } else {
                // get key from itemData and currentRangeConfig
                StringBuilder keyRowData = new StringBuilder();
                for (String columnItem : columnData) {
                    keyRowData.append(itemData.getValue(columnItem).toString());
                }

                if (!selectedDataTable.isExistKeyRowData(keyRowData.toString())) {
                    selectedDataTable.addKeyRowData(keyRowData.toString());

                    // Create new rowData
                    RowData newRowData = new RowData(itemData, keyRowData.toString(), level);
                    selectedDataTable.getRowData().add(newRowData);
                }

                // prepare param for recursive call
                LevelDataTable rLevelDataTable = selectedDataTable.getRowDataByKey(keyRowData.toString()).getLevelDataTable();
                int rLevel = level + 1;
                ArrayList<Range> rCurrentRangeConfig = selectedRangeConfig.getChildRange();

                // recursive call
                processDataRecursive(rLevelDataTable, itemData, rLevel, sheetConfig, rCurrentRangeConfig);
            }
        } else {
            // if current range config don't had index table excel => have only once element in ArrayList
            String[] columnData = currentRangeConfig.getFirst().getColumnData();

            // Check and process with no group column => leaf data
            if (columnData == null || columnData.length == 0 || (columnData.length == 1 && columnData[0].isEmpty())) {
                ArrayList<DataTable> dataTables = levelDataTable.getDataTables();

                // check if empty list, add new data table
                if (dataTables.isEmpty()) {
                    DataTable newDataTable = new DataTable();
                    RowData newRowData = new RowData(itemData, null, level);

                    newDataTable.getRowData().add(newRowData);

                    dataTables.add(newDataTable);
                } else {
                    // get exist data table and add new row data
                    DataTable selectedDataTable = dataTables.getFirst();

                    RowData newRowData = new RowData(itemData, null, level);

                    selectedDataTable.getRowData().add(newRowData);
                }
            } else {
                ArrayList<DataTable> dataTables = levelDataTable.getDataTables();

                // get key from itemData and currentRangeConfig
                StringBuilder keyRowData = new StringBuilder();
                for (String columnItem : columnData) {
                    keyRowData.append(itemData.getValue(columnItem).toString());
                }

                // check if empty list, add new data table
                if (dataTables.isEmpty()) {
                    DataTable newDataTable = new DataTable();
                    RowData newRowData = new RowData(itemData, keyRowData.toString(), level);

                    newDataTable.addKeyRowData(keyRowData.toString());
                    newDataTable.getRowData().add(newRowData);

                    dataTables.add(newDataTable);
                } else {
                    // get exist data table and add new row data
                    DataTable selectedDataTable = dataTables.getFirst();

                    // Check exist key
                    if (!selectedDataTable.isExistKeyRowData(keyRowData.toString())) {
                        selectedDataTable.addKeyRowData(keyRowData.toString());

                        RowData newRowData = new RowData(itemData, keyRowData.toString(), level);

                        selectedDataTable.getRowData().add(newRowData);
                    }
                }

                // prepare param for recursive call
                DataTable selectedDataTable = dataTables.getFirst();
                LevelDataTable rLevelDataTable = selectedDataTable.getRowDataByKey(keyRowData.toString()).getLevelDataTable();
                int rLevel = level + 1;
                ArrayList<Range> rCurrentRangeConfig = currentRangeConfig.getFirst().getChildRange();

                // recursive call
                processDataRecursive(rLevelDataTable, itemData, rLevel, sheetConfig, rCurrentRangeConfig);
            }
        }
    }

    private static void generateFile (XSSFWorkbook sourceTemplate, XSSFSheet targetSheet, ConfigSetting configSetting, SheetConfig sheetConfig, LevelDataTable rootLevelDataTable, JsonArray sourceData) throws Exception {
        long startTime;

        sourceTemplate.setForceFormulaRecalculation(true);
        XSSFFormulaEvaluator.evaluateAllFormulaCells(sourceTemplate);

        // Move footer
        System.out.print("\tMove footer... ");
        startTime = System.currentTimeMillis();
        int heightTable = calculateTotalTableHeightRecursive(sheetConfig, null, sheetConfig.getArrRange(), null, rootLevelDataTable.getDataTables(), 0);
        if (heightTable > 0) {
            targetSheet.shiftRows(new CellAddress(sheetConfig.getArrRange().getFirst().getEnd()).getRow() + 1, targetSheet.getLastRowNum(), heightTable, true, true);
        }
        System.out.println((System.currentTimeMillis() - startTime) + "ms");

        // Init start row
        int startRow = new CellAddress(sheetConfig.getArrRange().getFirst().getEnd()).getRow() + 1;
        if (heightTable > 0) {
            System.out.print("\tGenerate... ");
            startTime = System.currentTimeMillis();
            generateFileFromTemplate(startRow, targetSheet, sheetConfig, null, sheetConfig.getArrRange(), null, rootLevelDataTable.getDataTables(), 0);
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
            fillDataGeneral(targetSheet, heightTable, configSetting, sheetConfig, sourceData.getJsonObject(0));
            System.out.println((System.currentTimeMillis() - startTime) + "ms");
        }

        // Merge cell
        if (configSetting.isMergeCell() && heightTable > 0) {
            startTime = System.currentTimeMillis();
            System.out.print("\tMerge cell... ");
            System.out.println((System.currentTimeMillis() - startTime) + "ms");
            mergeCell(targetSheet);
        }

        XSSFFormulaEvaluator.evaluateAllFormulaCells(sourceTemplate);
    }

    private static int generateFileFromTemplate(int startRow, XSSFSheet targetSheet, SheetConfig sheetConfig, Range parentRange, List<Range> rangeList, DataTable parentDataTable, List<DataTable> dataTableList, int level) throws Exception {
        int totalAppendRow = 0;

        // loop all dataTable
        for (int indexRangeList = 0; indexRangeList < rangeList.size(); indexRangeList++) {
            Range rangeConfig = rangeList.get(indexRangeList);

            if (dataTableList.size() > indexRangeList) {
                DataTable selectedDataTable = dataTableList.get(indexRangeList);

                // loop all rowData in dataTable
                for (int indexRowData = 0; indexRowData < selectedDataTable.getRowData().size(); indexRowData++) {
                    RowData selectedRowData = selectedDataTable.getRowData().get(indexRowData);

                    // generate for highest row (leaf)
                    if (rangeConfig.getChildRange() == null) {
                        CellAddress beginTemplate = new CellAddress(rangeConfig.getBegin());
                        CellAddress endTemplate = new CellAddress(rangeConfig.getEnd());

                        int highRow = rangeConfig.getHeightRange();

//                        System.out.println("level: " + level + "    generate leaf           (" + (beginTemplate.getRow()) + "," + (endTemplate.getRow()) + ") -> " + (startRow + totalAppendRow));
                        targetSheet.copyRows(beginTemplate.getRow(), endTemplate.getRow(), startRow + totalAppendRow, new CellCopyPolicy());

                        // fill data ...
                        CellAddress beginCellAddress = new CellAddress(startRow + totalAppendRow, beginTemplate.getColumn());
                        CellAddress endCellAddress = new CellAddress(startRow + totalAppendRow + highRow, endTemplate.getColumn());
                        Range rangeFillData = new Range(beginCellAddress.toString(), endCellAddress.toString());
                        fillData(rangeFillData, selectedRowData.getRowData(), targetSheet, null, "<#table.(.*?)>");

                        totalAppendRow += highRow;
                    }

                    // generate for sub data row (branch)
                    if (rangeConfig.getChildRange() != null) {

                        CellAddress beginCellAddress = null;
                        CellAddress endCellAddress = null;

                        // generate begin to begin child
                        {
                            int beginRowTemplate = new CellAddress(rangeConfig.getBegin()).getRow();
                            int beginRowChildTemplate = new CellAddress(rangeConfig.getChildRange().getFirst().getBegin()).getRow();

                            int highRow = beginRowChildTemplate - beginRowTemplate;

                            beginCellAddress = new CellAddress(startRow + totalAppendRow, new CellAddress(rangeConfig.getBegin()).getColumn());

                            if (highRow > 0) {
//                                System.out.println("level: " + level + "    generate top            (" + (beginRowTemplate) + "," + (beginRowChildTemplate - 1) + ") -> " + (startRow + totalAppendRow));
                                targetSheet.copyRows(beginRowTemplate, beginRowChildTemplate - 1, startRow + totalAppendRow, new CellCopyPolicy());

                                totalAppendRow += highRow;
                            }

                        }

                        // recursive - generate child row
                        {
                            int childRowNum = generateFileFromTemplate(startRow + totalAppendRow, targetSheet, sheetConfig, rangeConfig, rangeConfig.getChildRange(), selectedDataTable, selectedRowData.getLevelDataTable().getDataTables(), level + 1);

                            totalAppendRow += childRowNum;
                        }

                        // generate end last child to end row
                        {
                            int endRowTemplate = new CellAddress(rangeConfig.getEnd()).getRow();
                            int endRowChildTemplate = new CellAddress(rangeConfig.getChildRange().getLast().getEnd()).getRow();

                            int highRow = endRowTemplate - endRowChildTemplate;

                            endCellAddress = new CellAddress(startRow + totalAppendRow + highRow, new CellAddress(rangeConfig.getEnd()).getColumn());

                            if (highRow > 0) {
//                                System.out.println("level: " + level + "    generate bottom         (" + (endRowChildTemplate + 1) + "," + (endRowTemplate) + ") -> " + (startRow + totalAppendRow));
                                targetSheet.copyRows(endRowChildTemplate + 1, endRowTemplate, startRow + totalAppendRow, new CellCopyPolicy());

                                totalAppendRow += highRow;
                            }
                        }

                        // fill data ...
                        Range rangeFillData = new Range(beginCellAddress.toString(), endCellAddress.toString());
                        fillData(rangeFillData, selectedRowData.getRowData(), targetSheet, null, "<#table.(.*?)>");
                    }
                }
            }

            // generate space between two datatable - skip last dataTable
            if (indexRangeList != rangeList.size() - 1) {
                Range nextRangeConfig = rangeList.get(indexRangeList + 1);

                int endRowRangeTemplate = new CellAddress(rangeConfig.getEnd()).getRow();
                int startNextRowRangeTemplate = new CellAddress(nextRangeConfig.getBegin()).getRow();

                int highRow = startNextRowRangeTemplate - endRowRangeTemplate - 1;

//                System.out.println("level: " + level + "    generate space          (" + (endRowRangeTemplate + 1) + "," + (startNextRowRangeTemplate - 1) + ") -> " + (startRow + totalAppendRow));
                targetSheet.copyRows(endRowRangeTemplate + 1, startNextRowRangeTemplate - 1, startRow + totalAppendRow, new CellCopyPolicy());
//                exportTempFile(targetSheet);

                totalAppendRow += highRow;
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
        int firstRowNum = new CellAddress(sheetConfig.getArrRange().getFirst().getBegin()).getRow();
        int lastRowNum = new CellAddress(sheetConfig.getArrRange().getFirst().getEnd()).getRow();
        int firstColNum = new CellAddress(sheetConfig.getArrRange().getFirst().getBegin()).getColumn();
        int lastColNum = new CellAddress(sheetConfig.getArrRange().getFirst().getEnd()).getColumn();

        // remove data
        for (int index = 0; index < heightParent; index++) {
            XSSFRow row = sheet.getRow(index + firstRowNum);
            sheet.removeRow(row);
        }

        // un-merge cell for shiftRow
        List<CellRangeAddress> listMergeCell = sheet.getMergedRegions();
        ArrayList<Integer> listIndexMergeRegions = new ArrayList<>();

        for (int i = 0; i < listMergeCell.size(); i++) {
            CellRangeAddress cellAddresses = listMergeCell.get(i);
            if (cellAddresses.getFirstRow() >= firstRowNum && cellAddresses.getLastRow() <= lastRowNum && cellAddresses.getFirstColumn() >= firstColNum && cellAddresses.getLastColumn() <= lastColNum) {
                listIndexMergeRegions.add(i);
            }
        }
        sheet.removeMergedRegions(listIndexMergeRegions);

        int lastRow = sheet.getLastRowNum();
        sheet.shiftRows(firstRowNum + heightParent, lastRow, -heightParent, true, true);
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

                    fillDataCellRecursive(cell, data, configSetting, regex);
                }
            }
        }
    }

    private static void fillDataCellRecursive(XSSFCell cell, JsonObject data, ConfigSetting configSetting, String regex) {
        if (cell == null) {
            return;
        }

        CellType a = cell.getCellType();

        String valueCell = "";
        try {
            valueCell = cell.getStringCellValue();
        } catch (Exception e) {
            return;
        }

        // fill data to cell
        if (valueCell != null && !valueCell.isEmpty()) {
            Pattern pattern = Pattern.compile(regex);
            Matcher matcher = pattern.matcher(valueCell);

            String key;
            if (matcher.find()) {
                key = matcher.group(1);
            } else {
                return;
            }

            String value = data.getString(key);
            if (value == null) {
                return;
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

            fillDataCellRecursive(cell, data, configSetting, regex);
        }
    }

    private static int calculateTotalTableHeightRecursive(SheetConfig sheetConfig, Range parentRange, List<Range> rangeList, DataTable parentDataTable, List<DataTable> dataTableList, int level) {
        int totalHeight = 0;

        int spaceBetweenTopParent = 0;
        int spaceBetweenBottomParent = 0;

        if (level != 0) {
            // calc space from begin parent to first child
            int beginParent = new CellAddress(parentRange.getBegin()).getRow();
            int beginFirstChild = new CellAddress(rangeList.getFirst().getBegin()).getRow();
            spaceBetweenTopParent = beginFirstChild - beginParent;

            // calc space from last child to end parent
            int endParent = new CellAddress(parentRange.getEnd()).getRow();
            int endLastChild = new CellAddress(rangeList.getLast().getEnd()).getRow();
            spaceBetweenBottomParent = endParent - endLastChild;
        }

        // calc space between - loop skip last element
        int totalSpaceBetweenChild = 0;
        for (int indexRange = 0; indexRange < rangeList.size() - 1; indexRange++) {
            Range firstRangeCob = rangeList.get(indexRange);
            Range secondRangeCob = rangeList.get(indexRange + 1);

            // calc diff
            int intRangeFirstCob = new CellAddress(firstRangeCob.getEnd()).getRow();
            int intRangeSecondCob = new CellAddress(secondRangeCob.getBegin()).getRow();

            totalSpaceBetweenChild += (intRangeSecondCob - intRangeFirstCob - 1);
        }

        // calc child claim space
        int totalSpaceBetweenRowAndTable = 0;
        for (int indexRange = 0; indexRange < rangeList.size(); indexRange++) {
            Range selectedRange = rangeList.get(indexRange);

            if (dataTableList.size() <= indexRange) {
                continue;
            }

            DataTable dataTable = dataTableList.get(indexRange);
            List<RowData> rowDataList = dataTable.getRowData();

            int totalClaimRowChild = 0;
            for (int indexRowChild = 0; indexRowChild < rowDataList.size(); indexRowChild++) {

                // Level row table is max of tree
                if (selectedRange.getChildRange() == null) {
                    totalClaimRowChild += selectedRange.getHeightRange();
                } else {
                    int temp = calculateTotalTableHeightRecursive(sheetConfig, selectedRange, selectedRange.getChildRange(), dataTable, dataTable.getRowData().get(indexRowChild).getLevelDataTable().getDataTables(), level + 1);
                    totalClaimRowChild += temp;
                }
            }

            totalSpaceBetweenRowAndTable += totalClaimRowChild;
        }

        totalHeight += (spaceBetweenTopParent + spaceBetweenBottomParent + totalSpaceBetweenChild + totalSpaceBetweenRowAndTable);

        return totalHeight;
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