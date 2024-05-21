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
import java.io.IOException;
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

        JsonObject jsonArrData = processData(configSetting, sourceData);

        String encodeTemp = jsonArrData.encodePrettily();
        System.out.println(encodeTemp);

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
            JsonObject itemData = sourceData.getJsonObject(index);

            processDataRecursive(resultData, itemData, 0, configSetting);
        }

        return resultData;
    }

    private static void processDataRecursive (JsonObject resultDataLevel, JsonObject itemData, int level, ConfigSetting configSetting) {
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
                    processDataRecursive(keyOb.getJsonObject("child") ,itemData, level + 1, configSetting);
                }
            } else {
                if (level + 1 < configSetting.getTotalGroup()) {
                    processDataRecursive(findKeyString.getJsonObject("child") ,itemData, level + 1, configSetting);
                }
            }
        } else {
            JsonObject dataObject = resultDataLevel.getJsonObject("data");
            JsonObject valueObject = new JsonObject();
            valueObject.put("value", itemData);
            dataObject.put(itemData.getString("ROW_NUM"), valueObject);
        }
    }

    private static Range appendTemplateFollowLevel (int level, int rowDestNum, XSSFSheet sheet, ConfigSetting configSetting) throws Exception {
        Range selectedRange = configSetting.getArrRange()[level];

        CellAddress beginCellAddress = new CellAddress(selectedRange.getBegin());
        CellAddress endCellAddress = new CellAddress(selectedRange.getEnd());

        int lastRow = sheet.getLastRowNum();

        // Move row for get more space
        sheet.shiftRows(rowDestNum, lastRow, endCellAddress.getRow() - beginCellAddress.getRow() + 1, true, true);




        // Duplicate template
        sheet.copyRows(beginCellAddress.getRow(), endCellAddress.getRow(), rowDestNum, new CellCopyPolicy());

        FileOutputStream fOut = new FileOutputStream("./temp.xlsx");
        sheet.getWorkbook().write(fOut);
        fOut.close();

        XSSFCell newBeginCell = sheet.getRow(rowDestNum).getCell(beginCellAddress.getColumn());
        XSSFCell newEndCell = sheet.getRow(rowDestNum + (endCellAddress.getRow() - beginCellAddress.getRow())).getCell(endCellAddress.getColumn());

        return new Range(newBeginCell.getAddress().toString(), newEndCell.getAddress().toString());
    }

    private static Range copyTemplateFollowLevel (int level, int rowDestNum, XSSFSheet sheet, ConfigSetting configSetting) {
        Range selectedRange = configSetting.getArrRange()[level];

        CellAddress beginCellAddress = new CellAddress(selectedRange.getBegin());
        CellAddress endCellAddress = new CellAddress(selectedRange.getEnd());

        XSSFCell newBeginCell = sheet.getRow(rowDestNum - beginCellAddress.getRow() + endCellAddress.getRow() - 1).getCell(beginCellAddress.getColumn());
        XSSFCell newEndCell = sheet.getRow(rowDestNum - 1).getCell(endCellAddress.getColumn());

        return new Range(newBeginCell.getAddress().toString(), newEndCell.getAddress().toString());
    }

    private static void generateFile (XSSFWorkbook sourceTemplate, ConfigSetting configSetting, JsonObject jsonArrData, String path) throws Exception {
        // Get sheet data
        XSSFSheet sheet = sourceTemplate.getSheetAt(0);

        int beginRowP = new CellAddress(configSetting.getArrRange()[0].getEnd()).getRow() + 1;

        generateTemplateAndFillData(0, beginRowP, jsonArrData, sheet, configSetting);

        // remove range template
        removeRangeTemplate(configSetting, sheet);

        FileOutputStream fOut = new FileOutputStream(path);
        sourceTemplate.write(fOut);
        fOut.close();
    }

    private static int generateTemplateAndFillData (int level, int beginRowStart, JsonObject jsonArrData, XSSFSheet targetSheet, ConfigSetting configSetting) throws Exception {
        // Get data object from jsonArrData
        JsonObject data = jsonArrData.getJsonObject("data");
        int beginRow = beginRowStart;
        int indexData = 0;
        int totalAppendRow = 0;

        for (Map.Entry<String, Object> item: data) {
            System.out.println("level: " + level + " - Key: " + item.getKey());

            int appendRow = 0;
            if (indexData == 0 && level > 0) {
                Range newCopyAddressRange = copyTemplateFollowLevel(level, beginRow, targetSheet, configSetting);

                // fill data to new address range
                fillData(newCopyAddressRange, ((JsonObject) item.getValue()).getJsonObject("value"), targetSheet, configSetting);

            } else {
                Range newAppendAddressRange = appendTemplateFollowLevel(level, beginRow + totalAppendRow, targetSheet, configSetting);
                appendRow = newAppendAddressRange.getHeightRange();

                // fill data to new address range
                fillData(newAppendAddressRange, ((JsonObject) item.getValue()).getJsonObject("value"), targetSheet, configSetting);
            }

            FileOutputStream fOut = new FileOutputStream("./temp.xlsx");
            targetSheet.getWorkbook().write(fOut);
            fOut.close();

            totalAppendRow += appendRow;

            if (level + 1 < configSetting.getTotalGroup()) {

                JsonObject childData = ((JsonObject) item.getValue()).getJsonObject("child");
                appendRow = generateTemplateAndFillData(level + 1, beginRow + totalAppendRow, childData, targetSheet, configSetting);

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
                if (valueCell != null) {
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
}