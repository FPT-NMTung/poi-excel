package org.example;

import io.vertx.core.json.JsonArray;
import io.vertx.core.json.JsonObject;
import model.ConfigSetting;
import model.Range;
import org.apache.commons.io.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileReader;

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

        JsonArray jsonArrData = processData(sourceData);

        generateFile(wb, configSetting, jsonArrData);

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

    private static JsonArray processData (ConfigSetting configSetting, JsonArray sourceData) {
        JsonObject resultData = new JsonObject();

        for (int index = 0; index < sourceData.size(); index++) {
            JsonObject itemData = sourceData.getJsonObject(index);

            processDataRecursive(resultData, itemData, 0, configSetting);
        }

        return null;
    }

    private static void processDataRecursive (JsonObject resultDataLevel, JsonObject itemData, int level, ConfigSetting configSetting) {
        // Condition break out recursive
        if (level >= configSetting.getTotalGroup()) {
            return;
        }

        // Get data follow config file
        String[] columnData = configSetting.getArrRange()[level].getColumnData();
        if (columnData != null) {
            StringBuilder key = new StringBuilder();
            for (String columnDatum : columnData) {
                key.append(itemData.getValue(columnDatum).toString());
            }

            // Check exist key in result object
            resultDataLevel.get
        }
    }

    private static void generateFile (XSSFWorkbook sourceTemplate, ConfigSetting configSetting, JsonArray jsonArrData) throws Exception {

    }
}