package org.example;

import converter.Converter;
import io.vertx.core.json.JsonArray;
import io.vertx.core.json.JsonObject;
import modelDocx.*;
import org.apache.commons.io.IOUtils;
import org.apache.poi.hssf.record.CommonObjectDataSubRecord;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Shape;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTxbxContent;

import java.io.*;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.*;

public class MainDocx {
    public static void main(String[] args) throws Exception {
        // Read template
        File templateFile = new File("GD_10 - copy.docx");
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
            ArrayList<TableData> dataProcessed = processData(sourceData, configSetting);
            System.out.println((System.currentTimeMillis() - startTime) + "ms");

            // Fill table data
            startTime = System.currentTimeMillis();
            System.out.print("Fill table data ... ");
            generateTable(doc, dataProcessed, configSetting);
            System.out.println((System.currentTimeMillis() - startTime) + "ms");
        }

        FileOutputStream fOut = new FileOutputStream("./result.docx");
        doc.write(fOut);
        fOut.close();

        System.out.println("Export done!");
    }

    private static ArrayList<TableData> processData(JsonArray sourceData, ConfigSetting configSetting) {
        ArrayList<TableData> tableDataList = new ArrayList<>();

        // Get list table from config
        ArrayList<TableConfig> tableConfigList = configSetting.getTableConfigs();
        for (TableConfig tableConfig: tableConfigList) {
            TableData tableData = new TableData(tableConfig.getTableName());
            tableDataList.add(tableData);
        }

        for (Object rowDataObj: sourceData) {
            JsonObject rowData = (JsonObject) rowDataObj;
            String nameTable = rowData.getString("NAME_TABLE");

            // Find index table config
            int indexTable = 0;
            for (int indexTableConfig = 0; indexTableConfig < tableConfigList.size(); indexTableConfig++) {
                if (tableConfigList.get(indexTableConfig).getTableName().equals(nameTable)) {
                    indexTable = indexTableConfig;
                    break;
                }
            }

            RowConfig rowConfig = tableConfigList.get(indexTable).getRowConfig();
            ArrayList<RowData> tableData = tableDataList.get(indexTable).getRows();

            processDataRecursive(tableData, rowData, configSetting, rowConfig);
        }

        return tableDataList;
    }

    private static void processDataRecursive(ArrayList<RowData> listRowData, JsonObject rowDataObj, ConfigSetting configSetting, RowConfig rowConfig) {
        String indexRow = rowConfig.getIndex();
        String indexValue = rowDataObj.getString(indexRow);

        // check exist value in array row data
        int iValue = -1;
        for (int index = 0; index < listRowData.size(); index++) {
            RowData rowData = listRowData.get(index);
            if (rowData.getData().getString(indexRow).equals(indexValue)) {
                iValue = index;
                break;
            }
        }

        if (iValue == -1) {
            RowData newRowData = new RowData(rowDataObj);

            if (rowConfig.getRowChildConfig() != null) {
                ArrayList<RowData> newRowDataChild = newRowData.getChildRow();
                RowConfig rowConfigChild = rowConfig.getRowChildConfig();

                processDataRecursive(newRowDataChild, rowDataObj, configSetting, rowConfigChild);
            }

            listRowData.add(newRowData);
        } else {
            RowData selectedRowData = listRowData.get(iValue);

            if (rowConfig.getRowChildConfig() != null) {
                ArrayList<RowData> newRowDataChild = selectedRowData.getChildRow();
                RowConfig rowConfigChild = rowConfig.getRowChildConfig();

                processDataRecursive(newRowDataChild, rowDataObj, configSetting, rowConfigChild);
            }
        }
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

    private static void fillDataGeneral(XWPFDocument doc, JsonObject data, ConfigSetting configSetting) throws Exception {
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

//    private static void fillDataToShape(List<XWPFRun> listRuns, String searchText, String replacement) {
//        TextSegment searchTextSegment;
//        while((searchTextSegment = paragraph.searchText(searchText, new PositionInParagraph(0, 0, 0))) != null) {
//            XWPFRun beginRun = paragraph.getRuns().get(searchTextSegment.getBeginRun());
//            String textInBeginRun = beginRun.getText(searchTextSegment.getBeginText());
//            String textBefore = textInBeginRun.substring(0, searchTextSegment.getBeginChar());
//
//            XWPFRun endRun = paragraph.getRuns().get(searchTextSegment.getEndRun());
//            String textInEndRun = endRun.getText(searchTextSegment.getEndText());
//            String textAfter = textInEndRun.substring(searchTextSegment.getEndChar() + 1);
//
//            if (searchTextSegment.getEndRun() == searchTextSegment.getBeginRun()) {
//                textInBeginRun = textBefore + replacement + textAfter; // if we have only one run, we need the text before, then the replacement, then the text after in that run
//            } else {
//                textInBeginRun = textBefore + replacement; // else we need the text before followed by the replacement in begin run
//                endRun.setText(textAfter, searchTextSegment.getEndText()); // and the text after in end run
//            }
//
//            beginRun.setText(textInBeginRun, searchTextSegment.getBeginText());
//
//            for (int runBetween = searchTextSegment.getEndRun() - 1; runBetween > searchTextSegment.getBeginRun(); runBetween--) {
//                paragraph.removeRun(runBetween); // remove not needed runs
//            }
//        }
//    }

    private static void generateTable(XWPFDocument doc, ArrayList<TableData> listTableData, ConfigSetting configSetting) throws Exception {
        List<XWPFTable> tables = doc.getTables();

        int indexTableDoc = 0;
        for (XWPFTable table : tables) {
            // get name table from config
            if (!table.getText().contains("<#TBG>")) {
                continue;
            }

            TableConfig tableConfig = configSetting.getTableConfigs().get(indexTableDoc);
            RowConfig rowConfig = tableConfig.getRowConfig();
            ArrayList<RowData> listRowData = listTableData.get(indexTableDoc).getRows();

            // fill data
            generateTableRowTable(doc, table, rowConfig.getEndRow() + 1, listRowData, rowConfig);

            // remove template row
            for (int cursorRowIndex = rowConfig.getStartRow(); cursorRowIndex <= rowConfig.getEndRow(); cursorRowIndex++) {
                table.removeRow(rowConfig.getStartRow());
            }

            // remove template <#TBG>
            table.removeRow(table.getRows().size() - 1);

            indexTableDoc++;
        }
    }

    private static void test(XWPFDocument doc) throws Exception {
        FileOutputStream fOut = new FileOutputStream("./test.docx");
        doc.write(fOut);
        fOut.close();
    }

    private static int generateTableRowTable(XWPFDocument doc, XWPFTable table, int startIndexRow, ArrayList<RowData> listRowData, RowConfig rowConfig) throws Exception {
        int totalAppend = 0;

        if (rowConfig.getRowChildConfig() == null) {
            for (RowData rowData : listRowData) {
                // copy row from template to table
                for (int cursorRow = rowConfig.getStartRow(); cursorRow <= rowConfig.getEndRow(); cursorRow++) {
                    XWPFTableRow selectedRow = table.getRow(cursorRow);
                    XWPFTableRow copiedRow = new XWPFTableRow((CTRow) selectedRow.getCtRow().copy(), table);
                    int newPos = totalAppend + startIndexRow;

                    // loop for fill data
                    for (CellConfig cellConfig: rowConfig.getMapCellConfig()) {
                        String nameCell = cellConfig.getName();
                        String valueCell = rowData.getData().getValue(cellConfig.getData()).toString();
                        String formatValueCell = formatField(valueCell, cellConfig.getFormat());

                        for (XWPFTableCell cell: copiedRow.getTableCells()) {
                            for (XWPFParagraph cellParagraph: cell.getParagraphs()) {
                                fillDataToParagraph(cellParagraph, "<#table." + nameCell + ">", formatValueCell);
                            }
                        }
                    }

                    table.addRow(copiedRow, newPos);

                    totalAppend++;
                }
            }
        } else {
            for (RowData rowData : listRowData) {
                // copy row from template to table
                for (int cursorRow = rowConfig.getStartRow(); cursorRow <= rowConfig.getRowChildConfig().getStartRow() - 1; cursorRow++) {
                    XWPFTableRow selectedRow = table.getRow(cursorRow);
                    XWPFTableRow copiedRow = new XWPFTableRow((CTRow) selectedRow.getCtRow().copy(), table);
                    int newPos = totalAppend + startIndexRow;

                    // loop for fill data
                    for (CellConfig cellConfig: rowConfig.getMapCellConfig()) {
                        String nameCell = cellConfig.getName();
                        String valueCell = rowData.getData().getValue(cellConfig.getData()).toString();
                        String formatValueCell = formatField(valueCell, cellConfig.getFormat());

                        for (XWPFTableCell cell: copiedRow.getTableCells()) {
                            for (XWPFParagraph cellParagraph: cell.getParagraphs()) {
                                fillDataToParagraph(cellParagraph, "<#table." + nameCell + ">", formatValueCell);
                            }
                        }
                    }

                    table.addRow(copiedRow, newPos);

                    totalAppend++;
                }

                test(doc);

                RowConfig childRowConfig = rowConfig.getRowChildConfig();
                ArrayList<RowData> listRowDataChild = rowData.getChildRow();
                totalAppend += generateTableRowTable(doc, table, totalAppend + startIndexRow, listRowDataChild, childRowConfig);

                for (int cursorRow = rowConfig.getRowChildConfig().getEndRow() + 1; cursorRow <= rowConfig.getEndRow(); cursorRow++) {
                    XWPFTableRow selectedRow = table.getRow(cursorRow);
                    XWPFTableRow copiedRow = new XWPFTableRow((CTRow) selectedRow.getCtRow().copy(), table);
                    int newPos = totalAppend + startIndexRow;

                    // loop for fill data
                    for (CellConfig cellConfig: rowConfig.getMapCellConfig()) {
                        String nameCell = cellConfig.getName();
                        String valueCell = rowData.getData().getValue(cellConfig.getData()).toString();
                        String formatValueCell = formatField(valueCell, cellConfig.getFormat());

                        for (XWPFTableCell cell: copiedRow.getTableCells()) {
                            for (XWPFParagraph cellParagraph: cell.getParagraphs()) {
                                fillDataToParagraph(cellParagraph, "<#table." + nameCell + ">", formatValueCell);
                            }
                        }
                    }

                    table.addRow(copiedRow, newPos);

                    totalAppend++;
                }
            }
        }

        return totalAppend;
    }

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
                result = Converter.numberToCharVi(bd.stripTrailingZeros().toPlainString()).trim();
                break;
            case "number_char_Vi":
                bd = new BigDecimal(value);
                result = Converter.numberToCharVi(bd.stripTrailingZeros().toPlainString()).trim();
                if (result.length() >= 1) {
                    result = result.substring(0, 1).toUpperCase() + result.substring(1);
                }
                break;
            case "number_char_VI":
                bd = new BigDecimal(value);
                result = Converter.numberToCharVi(bd.stripTrailingZeros().toPlainString()).trim().toUpperCase();
                break;
            case "number_char_en":
                result = Converter.numberToCharEn(value).trim();
                break;
            case "number_char_En":
                result = Converter.numberToCharEn(value).trim();
                if (result.length() >= 1) {
                    result = result.substring(0, 1).toUpperCase() + result.substring(1);
                }
                break;
            case "number_char_EN":
                result = Converter.numberToCharEn(value).trim().toUpperCase();
                break;
            default:
                result = value;
        }

        return result;
    }
}
