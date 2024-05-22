package org.example;

import io.vertx.core.json.JsonArray;
import io.vertx.core.json.JsonObject;
import model.ConfigSetting;
import model.Range;
import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
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

        System.out.println("Start process data...");
        JsonObject jsonArrData = processData(configSetting, sourceData);

        System.out.println("Start generate data...");
        generateFile(wb, configSetting, jsonArrData, "./result.xlsx");

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

        ConfigSetting configSetting = new ConfigSetting(totalGroup);
        Range[] ranges = configSetting.getArrRange();

        // Get array object range
        int totalRow = sheet.getPhysicalNumberOfRows();
        int countRow = 0;
        int count = 0;
        while (count < totalGroup || countRow < 30) {
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
            System.out.println("data index: " + index);
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

    private static Range appendTemplateFollowLevel (int level, int startRow, int endRow, XSSFSheet sheet, ConfigSetting configSetting) throws Exception {
        Range selectedRange = configSetting.getArrRange()[level];

        CellAddress beginCellAddress = new CellAddress(selectedRange.getBegin());
        CellAddress endCellAddress = new CellAddress(selectedRange.getEnd());

        int lastRow = sheet.getLastRowNum();

        // Move row for get more space
        int nextEndRow = endRow + 1;
        int heightTemplateLevel = selectedRange.getHeightRange();
//        sheet.shiftRows(nextEndRow, lastRow, heightTemplateLevel, true, true);

        // Duplicate template
        int beginRowCopy = beginCellAddress.getRow();
        int endRowCopy = endCellAddress.getRow();
        sheet.copyRows(beginRowCopy, endRowCopy, nextEndRow, new CellCopyPolicy());

        FileOutputStream fOut = new FileOutputStream("./temp.xlsx");
        sheet.getWorkbook().write(fOut);
        fOut.close();

        // Create new address range
        int colNumBeginCellAddress = beginCellAddress.getColumn();
        int colNumEndCellAddress = endCellAddress.getColumn();
        CellAddress newBeginCell = new CellAddress(nextEndRow, colNumBeginCellAddress);
        CellAddress newEndCell = new CellAddress(endRow + heightTemplateLevel, colNumEndCellAddress);

        return new Range(newBeginCell.toString(), newEndCell.toString());
    }

    private static Range copyTemplateFollowLevel (int level, int startRow, int endRow, XSSFSheet sheet, ConfigSetting configSetting) {
        Range selectedRange = configSetting.getArrRange()[level];
        Range selectedRangeParent = configSetting.getArrRange()[level - 1];

        CellAddress beginCellAddress = new CellAddress(selectedRange.getBegin());
        CellAddress endCellAddress = new CellAddress(selectedRange.getEnd());

        CellAddress beginCellAddressParent = new CellAddress(selectedRangeParent.getBegin());

        int rowNumBeginCellAddress = startRow + (beginCellAddress.getRow() - beginCellAddressParent.getRow());
        int colNumBeginCellAddress = beginCellAddress.getColumn();
        int colNumEndCellAddress = endCellAddress.getColumn();
        int heightTemplateLevel = selectedRange.getHeightRange();

        CellAddress newBeginCell = new CellAddress(rowNumBeginCellAddress, colNumBeginCellAddress);
        CellAddress newEndCell = new CellAddress(rowNumBeginCellAddress + heightTemplateLevel - 1, colNumEndCellAddress);

        return new Range(newBeginCell.toString(), newEndCell.toString());
    }

    private static void generateFile (XSSFWorkbook sourceTemplate, ConfigSetting configSetting, JsonObject jsonArrData, String path) throws Exception {
        // Get sheet data
        XSSFSheet sheet = sourceTemplate.getSheetAt(0);

        int endRowTemplate = new CellAddress(configSetting.getArrRange()[0].getEnd()).getRow();

        // Move footer
        int heightTable = calculateTotalTableHeightRecursive(configSetting, jsonArrData, 0);
        sheet.shiftRows(new CellAddress(configSetting.getArrRange()[0].getEnd()).getRow() + 1, sheet.getLastRowNum(), heightTable, true, true);

        generateTemplateAndFillData(0, 0, endRowTemplate, jsonArrData, sheet, configSetting);

        // remove range template
//        removeRangeTemplate(configSetting, sheet);

        FileOutputStream fOut = new FileOutputStream(path);
        sourceTemplate.write(fOut);
        fOut.close();
    }

    private static int generateTemplateAndFillData (int level, int startRowParent, int endRowParent, JsonObject jsonArrData, XSSFSheet targetSheet, ConfigSetting configSetting) throws Exception {
        // Get data object from jsonArrData
        JsonObject data = jsonArrData.getJsonObject("data");
        int endRow = endRowParent;
        int startRow = startRowParent;
        int indexData = 0;
        int totalAppendRow = 0;

        for (Map.Entry<String, Object> item: data) {
            System.out.println("level: " + level + " - Key: " + item.getKey());

            int appendRow = 0;
            if (indexData == 0 && level > 0) {
                Range newCopyAddressRange = copyTemplateFollowLevel(level, startRow, endRow + totalAppendRow, targetSheet, configSetting);
                startRow = new CellAddress(newCopyAddressRange.getBegin()).getRow();

                // fill data to new address range
                fillData(newCopyAddressRange, ((JsonObject) item.getValue()).getJsonObject("value"), targetSheet, configSetting);
            } else {
                Range newAppendAddressRange = appendTemplateFollowLevel(level, startRow, endRow + totalAppendRow, targetSheet, configSetting);
                appendRow = newAppendAddressRange.getHeightRange();
                startRow = new CellAddress(newAppendAddressRange.getBegin()).getRow();

                // fill data to new address range
                fillData(newAppendAddressRange, ((JsonObject) item.getValue()).getJsonObject("value"), targetSheet, configSetting);
            }

            FileOutputStream fOut = new FileOutputStream("./temp.xlsx");
            targetSheet.getWorkbook().write(fOut);
            fOut.close();

            totalAppendRow += appendRow;

            if (level + 1 < configSetting.getTotalGroup()) {

                JsonObject childData = ((JsonObject) item.getValue()).getJsonObject("child");
                appendRow = generateTemplateAndFillData(level + 1, startRow, endRow + totalAppendRow, childData, targetSheet, configSetting);

                totalAppendRow += appendRow;
            }

            indexData += 1;
        }

        return totalAppendRow;
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

                String valueCell = cell.getStringCellValue();
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

                    cell.setCellValue(newValue);
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
}