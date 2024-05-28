package org.example;

import converter.Converter;
import io.vertx.core.json.JsonArray;
import io.vertx.core.json.JsonObject;
import modelDocx.TableData;
import org.apache.commons.io.IOUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTComments;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
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

        // Get config setting
        JsonObject configSetting = getConfigSetting(doc);

        // Process data
        startTime = System.currentTimeMillis();
        System.out.print("Process data ... ");
        HashMap<String, TableData> dataProcessed = processData(sourceData, configSetting);
        System.out.println((System.currentTimeMillis() - startTime) + "ms");

        // Fill general data
        startTime = System.currentTimeMillis();
        System.out.print("Fill general data ... ");
        fillDataGeneral(doc, sourceData.getJsonObject(0), configSetting);
        System.out.println((System.currentTimeMillis() - startTime) + "ms");

        // Fill table data
        startTime = System.currentTimeMillis();
        System.out.print("Fill table data ... ");
        fillDataTable(doc, dataProcessed, configSetting);
        System.out.println((System.currentTimeMillis() - startTime) + "ms");

        FileOutputStream fOut = new FileOutputStream("./result.docx");
        doc.write(fOut);
        fOut.close();

        System.out.println("Export done!");
    }

    private static HashMap<String, TableData> processData(JsonArray sourceData, JsonObject configSetting) {
        HashMap<String, TableData> tableDataList = new HashMap<>();

        // Get list table from config
        JsonObject listTableConfig = configSetting.getJsonObject("table");
        for (Map.Entry<String, Object> tableConfigObj: listTableConfig) {
            String nameTable = tableConfigObj.getKey();

            // check exist
            TableData tableData;
            if (tableDataList.containsKey(nameTable)) {
                tableDataList.get(nameTable);
            } else {
                tableData = new TableData(nameTable);
                tableDataList.put(nameTable, tableData);
            }
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

    private static JsonObject getConfigSetting(XWPFDocument doc) throws Exception {
        XWPFComments comments = doc.getDocComments();
        JsonObject jsonObject = new JsonObject();

        for (XWPFComment comment: comments.getComments()) {
            String content = comment.getText();

            jsonObject = new JsonObject(content);
        }

        if (Objects.equals(jsonObject.toString(), "{}")) {
            throw new Exception("Not found JSON config comment");
        }

        return jsonObject;
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

    private static void fillDataGeneral(XWPFDocument doc, JsonObject data, JsonObject configSetting) throws Exception {
        if (configSetting.getJsonArray("general") == null) {
            return;
        }

        // get list config general field
        JsonArray listFields = configSetting.getJsonArray("general");
        for (Object jsonObject: listFields) {

            String nameData = ((JsonObject) jsonObject).getString("name");
            String colData = ((JsonObject) jsonObject).getString("data");
            String format = ((JsonObject) jsonObject).getString("format");

            if (colData == null || colData.isEmpty()) {
                colData = nameData;
            }

            String searchText = "<#general." + nameData + ">";
            String replacement = data.getString(colData);

            replacement = formatField(replacement, format);

            // search all paragraph

            for (XWPFParagraph paragraph : doc.getParagraphs()) {
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
        }
    }

    private static void fillDataTable(XWPFDocument doc, HashMap<String, TableData> data, JsonObject configSetting) throws Exception {
        if (configSetting.getJsonObject("table") == null) {
            return;
        }

        List<XWPFTable> tables = doc.getTables();
        int indexTable = 0;
        for (XWPFTable table : tables) {
            // get name table from config
            if (!table.getText().contains("<#TBG>")) {
                continue;
            }

            int indexTableConfig = 0;
            String nameTable = "";
            JsonObject configTable = new JsonObject();
            for (Map.Entry<String, Object> tableConfig: configSetting.getJsonObject("table")) {
                nameTable = tableConfig.getKey();
                configTable = (JsonObject) tableConfig.getValue();
                if (indexTableConfig == indexTable) {
                    break;
                }
            }
            indexTable++;

            // get data
            TableData tableData = data.get(nameTable);
            if (tableData == null) {
                continue;
            }

            for (JsonObject rowData: tableData.getRows()) {
                int startIndexTable = configTable.getInteger("start");
                XWPFTableRow row = table.getRow(startIndexTable);

                JsonArray columnConfig = configTable.getJsonArray("column");
                int indexCell = 0;
                for (XWPFTableCell cell: row.getTableCells()) {
                    List<XWPFParagraph> paragraphs = cell.getParagraphs();

                    for (XWPFParagraph paragraph: paragraphs) {
                        while (!paragraph.runsIsEmpty()) {
                            paragraph.removeRun(0);
                        }
                    }

                    JsonObject dataNameColumn = columnConfig.getJsonObject(indexCell);

                    String valueField = rowData.getString(dataNameColumn.getString("name"));
                    String format = dataNameColumn.getString("format");
                    cell.setText(formatField(valueField, format));
                    indexCell++;
                }

                table.addRow(row);
            }

            table.removeRow(configTable.getInteger("start"));
        }
    }

    private static String formatField(String value, String format) {
        String result = "";

        if (format == null) {
            return value;
        }

        switch (format) {
            case "number":
                String pattern = "";
                if (value.contains(".")) {
                    int lengthDiv = value.substring(value.indexOf(".")).length();
                    String[] listZero = new String[lengthDiv];
                    Arrays.fill(listZero, "0");

                    String subDiv = String.join("", listZero);
                    pattern = "#,###" + "." + subDiv;
                } else {
                    pattern = "#,###";
                }

                DecimalFormat formatter = new DecimalFormat(pattern);
                result = formatter.format(Double.parseDouble(value));
                break;
            case "number_char_vi":
                result = Converter.numberToCharVi(value);
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
