package org.example;

import converter.Converter;
import io.vertx.core.json.JsonArray;
import io.vertx.core.json.JsonObject;
import modelDocx.*;
import org.apache.commons.io.IOUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTComments;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.util.*;

public class MainDocx {
    public static void main(String[] args) throws Exception {
        // Read template
        File templateFile = new File("Mau 33C-THQ.docx");
        if (!templateFile.exists()) {
            throw new Exception("Template file not found");
        }

        XWPFDocument doc = new XWPFDocument(OPCPackage.open(templateFile));

        long startTime;

        // Get data
        String jsonStr = IOUtils.toString(new FileReader("./dataDocx.json"));
        JsonArray sourceData = new JsonArray(jsonStr);

        System.out.println("Data size: " + sourceData.size() + " |" + (sourceData.size() >= 5000 ? " Warning: SLOWWW": ""));

        // Get config setting
        ConfigSetting configSetting = getConfigSetting(doc);

        if (!configSetting.getGeneralData().isEmpty()) {
            // Fill general data
            startTime = System.currentTimeMillis();
            System.out.print("Fill general data ... ");
            fillDataGeneral(doc, sourceData.getJsonObject(0), configSetting);
            System.out.println((System.currentTimeMillis() - startTime) + "ms");
        }

        if (!configSetting.getTableConfigs().isEmpty()) {
            // Process data
            startTime = System.currentTimeMillis();
            System.out.print("Process data ... ");
            HashMap<String, TableData> dataProcessed = processData(sourceData, configSetting);
            System.out.println((System.currentTimeMillis() - startTime) + "ms");

            // Fill table data
//            startTime = System.currentTimeMillis();
//            System.out.print("Fill table data ... ");
//            fillDataTable(doc, dataProcessed, configSetting);
//            System.out.println((System.currentTimeMillis() - startTime) + "ms");
        }

        FileOutputStream fOut = new FileOutputStream("./result.docx");
        doc.write(fOut);
        fOut.close();

        System.out.println("Export done!");
    }

    private static HashMap<String, TableData> processData(JsonArray sourceData, ConfigSetting configSetting) {
        HashMap<String, TableData> tableDataList = new HashMap<>();

        // Get list table from config
        ArrayList<TableConfig> listTableConfig = configSetting.getTableConfigs();
        for (TableConfig tableConfig: listTableConfig) {
            TableData tableData = new TableData(tableConfig.getTableName());
            tableDataList.put(tableConfig.getTableName(), tableData);
        }

        for (Object rowDataObj: sourceData) {
            JsonObject rowData = (JsonObject) rowDataObj;
            String nameTable = rowData.getString("NAME_TABLE");

            TableData tableData = tableDataList.get(nameTable);

            String indexTable = rowData.getString(configSetting.getJsonObject("table").getJsonObject(nameTable).getString("index"));
            if (!tableData.getKey().contains(indexTable)) {
                tableData.getRows().add(rowData);
                tableData.getKey().add(indexTable);
            }
        }

        return tableDataList;
    }

    private static ConfigSetting getConfigSetting(XWPFDocument doc) throws Exception {
        XWPFComments comments = doc.getDocComments();
        ConfigSetting configSetting = new ConfigSetting();

        boolean isHasConfigComment = false;

        if (comments == null) {
            return configSetting;
        }

        for (XWPFComment comment: comments.getComments()) {
            String content = comment.getText();

            if (content == null || content.isEmpty()) {
                continue;
            }

            // parse content comment to json
            JsonObject jsonConfig = new JsonObject(content);

            // get config general
            getGeneralConfig(configSetting, jsonConfig);

            // get config table
            getConfigSettingTable(configSetting, jsonConfig);
            isHasConfigComment = true;
        }

        if (configSetting.getGeneralData().isEmpty() && configSetting.getTableConfigs().isEmpty() && isHasConfigComment) {
            throw new Exception("Not found JSON config comment");
        }

        removeAllComment(doc);

        return configSetting;
    }

    private static void getGeneralConfig(ConfigSetting configSetting, JsonObject jsonConfig) {
        // check hash general config
        if (!jsonConfig.containsKey("general")) {
            return;
        }

        JsonArray arrConfigGeneral = jsonConfig.getJsonArray("general");
        for (int i = 0; i < arrConfigGeneral.size(); i++) {
            JsonObject general = arrConfigGeneral.getJsonObject(i);

            String name = general.getString("name");
            String data = general.getString("data");
            String format = general.getString("format");

            CellConfig cellConfig = new CellConfig(name, data, format);

            configSetting.getGeneralData().add(cellConfig);
        }
    }

    private static void getConfigSettingTable(ConfigSetting configSetting, JsonObject jsonConfig) {
        // check hash table config
        if (!jsonConfig.containsKey("table")) {
            return;
        }

        JsonArray arrConfigTable = jsonConfig.getJsonArray("table");
        ArrayList<TableConfig> tableDataList = new ArrayList<>();

        for (int i = 0; i < arrConfigTable.size(); i++) {
            JsonObject tableObj = arrConfigTable.getJsonObject(i);

            String name = tableObj.getString("name");

            TableConfig tableConfig = new TableConfig();
            tableConfig.setTableName(name);

            JsonObject rowObj = tableObj.getJsonObject("row");
            RowConfig rowConfig = getConfigSettingRow(rowObj);

            tableConfig.setRowConfig(rowConfig);
            tableDataList.add(tableConfig);
        }

        configSetting.setTableConfigs(tableDataList);
    }

    private static RowConfig getConfigSettingRow(JsonObject rowObj) {
        String range = rowObj.getString("range");
        int begin = Integer.parseInt(range.split("\\|")[0]);
        int end = Integer.parseInt(range.split("\\|")[1]);
        String indexRow = rowObj.getString("index");

        RowConfig rowConfig = new RowConfig(indexRow, begin, end);

        JsonArray colObj = rowObj.getJsonArray("column");
        for (int j = 0; j < colObj.size(); j++) {
            JsonObject colObjObj = colObj.getJsonObject(j);

            String colName = colObjObj.getString("name");
            String colData = colObjObj.getString("data");
            String colFormat = colObjObj.getString("format");

            CellConfig cellConfig = new CellConfig(colName, colData, colFormat);

            rowConfig.getMapCellConfig().add(cellConfig);
        }

        // Check deep table
        if (rowObj.containsKey("row")) {
            JsonObject childRowObj = rowObj.getJsonObject("row");
            RowConfig childRowConfig = getConfigSettingRow(childRowObj);

            rowConfig.setRowChildConfig(childRowConfig);
        }

        return rowConfig;
    }

    private static void removeAllComment(XWPFDocument doc) {
        for (XWPFParagraph paragraph : doc.getParagraphs()) {
            // remove all comment range start marks
            for (int i = paragraph.getCTP().getCommentRangeStartList().size() - 1; i >= 0; i--) {
                paragraph.getCTP().removeCommentRangeStart(i);
            }
            // remove all comment range end marks
            for (int i = paragraph.getCTP().getCommentRangeEndList().size() - 1; i >= 0; i--) {
                paragraph.getCTP().removeCommentRangeEnd(i);
            }
            // remove all comment references
            for (int i = paragraph.getRuns().size() - 1; i >= 0; i--) {
                XWPFRun run = paragraph.getRuns().get(i);
                if (!run.getCTR().getCommentReferenceList().isEmpty()) {
                    paragraph.removeRun(i);
                }
            }
        }
    }

    private static void fillDataGeneral(XWPFDocument doc, JsonObject data, ConfigSetting configSetting) {
        // get list config general field
        ArrayList<CellConfig> listCell = configSetting.getGeneralData();
        for (CellConfig cellConfig: listCell) {

            String nameData = cellConfig.getName();
            String colData = cellConfig.getData();
            String format = cellConfig.getFormat();

            if (colData == null || colData.isEmpty()) {
                colData = nameData;
            }

            String searchText = "<#general." + nameData + ">";
            String replacement = data.getString(colData);

            replacement = formatField(replacement, format);

            // search all paragraph

            for (XWPFParagraph paragraph : doc.getParagraphs()) {
                fillDataToParagraph(paragraph, searchText, replacement);
            }

            List<XWPFTable> tables = doc.getTables();
            for (XWPFTable table : tables) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            fillDataToParagraph(paragraph, searchText, replacement);
                        }
                    }
                }
            }
        }
    }

    private static void fillDataToParagraph(XWPFParagraph paragraph, String searchText, String replacement) {
        TextSegment searchTextSegment;
        while((searchTextSegment = paragraph.searchText(searchText, new PositionInParagraph(0, 0, 0))) != null) {
            XWPFRun beginRun = paragraph.getRuns().get(searchTextSegment.getBeginRun());
            String textInBeginRun = beginRun.getText(searchTextSegment.getBeginText());
            String textBefore = textInBeginRun.substring(0, searchTextSegment.getBeginChar());

            XWPFRun endRun = paragraph.getRuns().get(searchTextSegment.getEndRun());
            String textInEndRun = endRun.getText(searchTextSegment.getEndText());
            String textAfter = textInEndRun.substring(searchTextSegment.getEndChar() + 1);

            if (searchTextSegment.getEndRun() == searchTextSegment.getBeginRun()) {
                textInBeginRun = textBefore + replacement + textAfter; // if we have only one run, we need the text before, then the replacement, then the text after in that run
            } else {
                textInBeginRun = textBefore + replacement; // else we need the text before followed by the replacement in begin run
                endRun.setText(textAfter, searchTextSegment.getEndText()); // and the text after in end run
            }

            beginRun.setText(textInBeginRun, searchTextSegment.getBeginText());

            for (int runBetween = searchTextSegment.getEndRun() - 1; runBetween > searchTextSegment.getBeginRun(); runBetween--) {
                paragraph.removeRun(runBetween); // remove not needed runs
            }
        }
    }

//    private static void fillDataTable(XWPFDocument doc, HashMap<String, TableData> data, JsonObject configSetting) throws Exception {
//        if (configSetting.getJsonObject("table") == null) {
//            return;
//        }
//
//        List<XWPFTable> tables = doc.getTables();
//        int indexTable = 0;
//        for (XWPFTable table : tables) {
//            // get name table from config
//            if (!table.getText().contains("<#TBG>")) {
//                continue;
//            }
//
//            int indexTableConfig = 0;
//            String nameTable = "";
//            JsonObject configTable = new JsonObject();
//            for (Map.Entry<String, Object> tableConfig: configSetting.getJsonObject("table")) {
//                nameTable = tableConfig.getKey();
//                configTable = (JsonObject) tableConfig.getValue();
//                if (indexTableConfig == indexTable) {
//                    break;
//                }
//                indexTableConfig++;
//            }
//            indexTable++;
//
//            // get data
//            TableData tableData = data.get(nameTable);
//            if (tableData == null) {
//                continue;
//            }
//
//            int indexRowData = 0;
//            for (JsonObject rowData: tableData.getRows()) {
//                int startIndexTable = configTable.getInteger("start");
//                XWPFTableRow row = table.getRow(startIndexTable);
//
//                JsonArray columnConfig = configTable.getJsonArray("column");
//                int indexCell = 0;
//                for (XWPFTableCell cell: row.getTableCells()) {
//                    List<XWPFParagraph> paragraphs = cell.getParagraphs();
//
//                    for (XWPFParagraph paragraph: paragraphs) {
//                        while (!paragraph.runsIsEmpty()) {
//                            paragraph.removeRun(0);
//                        }
//                    }
//
//                    JsonObject dataNameColumn = columnConfig.getJsonObject(indexCell);
//
//                    String valueField = rowData.getString(dataNameColumn.getString("name"));
//                    String format = dataNameColumn.getString("format");
//                    cell.setText(formatField(valueField, format));
//                    indexCell++;
//                }
//
//                table.addRow(row, configTable.getInteger("start") + indexRowData + 1);
//                indexRowData++;
//            }
//
//            table.removeRow(configTable.getInteger("start"));
//        }
//    }

    private static String formatField(String value, String format) {
        String result = "";
        BigDecimal bd;
        if (format == null) {
            return value;
        }

        switch (format) {
            case "number":
                bd  = new BigDecimal(value);
                String newValue = bd.stripTrailingZeros().toPlainString();

                String pattern = "";
                if (newValue.contains(".")) {
                    int lengthDiv = newValue.substring(newValue.indexOf(".")).length();
                    String[] listZero = new String[lengthDiv];
                    Arrays.fill(listZero, "0");

                    String subDiv = String.join("", listZero);
                    pattern = "#,###" + "." + subDiv;
                } else {
                    pattern = "#,###";
                }

                DecimalFormat formatter = new DecimalFormat(pattern);
                result = formatter.format(Double.parseDouble(newValue));
                break;
            case "number_char_vi":
                bd = new BigDecimal(value);
                result = Converter.numberToCharVi(bd.stripTrailingZeros().toPlainString());
                break;
            case "number_char_en":
                result = Converter.numberToCharEn(value);
                break;
            default:
                result = value;
        }

        return result;
    }
}
