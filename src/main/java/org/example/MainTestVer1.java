package org.example;

import io.vertx.core.json.JsonArray;
import io.vertx.core.json.JsonObject;
import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.*;

public class MainTestVer1 {
    public static void main(String[] args) throws Exception {
        // Get file template
        File templateFile = new File("DE164.xlsx");
//        File templateFile = new File("GD_07.xlsx");

        if (!templateFile.exists()) {
            throw new Exception("Template file not found");
        }

        long startTime;
        long beginTime = System.currentTimeMillis();

        // Convert to POI
        XSSFWorkbook wb = new XSSFWorkbook(templateFile);

        // Get JSON data
        String jsonStr = IOUtils.toString(new FileReader("./data.json"));
        JsonArray sourceData = new JsonArray(jsonStr);

        generateExcel(wb, sourceData);


        FileOutputStream fOut = new FileOutputStream("./result.xlsx");
        wb.write(fOut);
        fOut.close();
    }

    private static final BorderStyle DEFAULT_BORDER_STYLE = BorderStyle.THIN;
    private static final int DEFAULT_MAX_WIDTH = 1024;

    public static XSSFWorkbook generateExcel(
            XSSFWorkbook xssfWorkbook, JsonArray jsonArray) throws Exception {
        if (Objects.isNull(jsonArray)) {
            return null;
        }
        XSSFWorkbook workbook = xssfWorkbook;
        XSSFSheet oldSheet = xssfWorkbook.getSheetAt(0);

//        if (Objects.nonNull(xssfWorkbook)) {
//            oldSheet = xssfWorkbook.getSheetAt(0);
//            workbook = new SXSSFWorkbook(xssfWorkbook);
////            workbook.setCompressTempFiles(true);
//        } else {
//            workbook = xssfWorkbook;
//        }

        XSSFSheet sheet = workbook.getSheetAt(0);

        int firstRowData = findStartRowIndex(oldSheet) - 1;
        int endRowData = firstRowData + jsonArray.size();
        fillData(jsonArray, workbook, oldSheet);

        long start = System.currentTimeMillis();
        Row rowDataTemplate = oldSheet.getRow(firstRowData);
        Collection<Integer> numberCollection = new ArrayList<>();
        for (int i = 0; i <= 25; i++) {
            numberCollection.add(i);
        }
//        sheet.trackColumnsForAutoSizing(numberCollection);
//        int rowIndex = firstRowData;
//        for (int i = 0; i <= rowDataTemplate.getPhysicalNumberOfCells(); i++) {
//            xssfWorkbook.getSheetAt(0).autoSizeColumn(i);
//        }

        return workbook;
    }

    //    public static SXSSFWorkbook generateExcel(
//            XSSFWorkbook xssfWorkbook, JsonObject jsonOject) {
//        if (jsonOject.size() == 0) {
//            return null;
//        }
//
//        SXSSFWorkbook workbook = null;
//        XSSFSheet oldSheet = null;
//
//        if (Objects.nonNull(xssfWorkbook)) {
//            oldSheet = xssfWorkbook.getSheetAt(0);
//            workbook = new SXSSFWorkbook(xssfWorkbook);
//            workbook.setCompressTempFiles(true);
//        } else {
//            workbook = new SXSSFWorkbook();
//        }
//
//        SXSSFSheet sheet = workbook.getSheetAt(0);
//
//        int firstRowData = findStartRowIndex(oldSheet) - 1;
//        int endRowData = firstRowData + jsonArray.size();
//        ExcelUtils.fillData(jsonArray, workbook, oldSheet);
//
//        long start = System.currentTimeMillis();
//        Row rowDataTemplate = oldSheet.getRow(firstRowData);
//        Collection<Integer> numberCollection = new ArrayList<>();
//        for (int i = 0; i <= 12; i++) {
//            numberCollection.add(i);
//        }
//        sheet.trackColumnsForAutoSizing(numberCollection);
//        int rowIndex = firstRowData;
//        for (int i = 0; i <= rowDataTemplate.getPhysicalNumberOfCells(); i++) {
//            xssfWorkbook.getSheetAt(0).autoSizeColumn(i);
//        }
//        Main.LOGGER.info("Autosize column take " + (System.currentTimeMillis() - start) + " ms.");
//        setFitWidthPage(CommonUtils.getOrDefault(oldSheet, sheet));
//        return workbook;
//    }
    private static void setFitWidthPage(Sheet sheet) {
        long start = System.currentTimeMillis();
        sheet.setFitToPage(true);
        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setFitWidth((short) 1);
        printSetup.setFitHeight((short) 0);
    }

    private static int calculateDesiredColumnWidth(XSSFWorkbook xssfWorkbook, int startRow, int endRow, Sheet sheet, int columnIndex) {
        int maxColumnWidth = 0;
        int maxSheet = 0;
        Row row;
        Sheet s = xssfWorkbook.getSheetAt(0);
        row = s.getRow(startRow);
        if (row != null) {
            Cell cell = row.getCell(columnIndex);
            if (cell != null) {
                int cellWidth = cell.toString().length();
                maxColumnWidth = Math.max(maxColumnWidth, cellWidth);
            }
        }
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
            row = s.getRow(rowIndex);
            if (row != null) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null) {
                    int cellWidth = cell.toString().length();
                    if (cellWidth > maxColumnWidth) {
                        maxSheet = 1;
                    }
                    maxColumnWidth = Math.max(maxColumnWidth, cellWidth);
                }
            }
        }
        return maxSheet;
    }

    private static void removeRowIfExists(XSSFSheet sheet, int index, int totalRow) {
        for (int i = 0; i < totalRow; i++) {
            Row row = sheet.getRow(index + i);
            if (Objects.nonNull(row)) {
                sheet.removeRow(row);
            }
        }
    }

    private static void fillData(
            JsonArray jsonArray,
            Workbook workbook,
            XSSFSheet sheet) throws Exception {

        int arraySize = jsonArray.size();
        CellStyle valueCellStyle = createValueCellStyle(workbook);

        if (arraySize > 0) {
            Row row;
            Cell cell;

            int columnNum = 0;
            long start = System.currentTimeMillis();

            int lastRowTemplate = sheet.getLastRowNum();
            int firstRowData = findStartRowIndex(sheet);
            Row rowDataTemplate = sheet.getRow(firstRowData);
            int indexRow = firstRowData;
            // Di chuyển các dòng bằng cách sử dụng phương thức shiftRows()
            sheet.shiftRows(firstRowData + 1, lastRowTemplate, arraySize, true, true);

            // fill data
            for (int i = 0; i < arraySize; i++) {
                JsonObject jsonObject = jsonArray.getJsonObject(i);
                row = getOrCreateRow(sheet, ++indexRow);
                Set<Map.Entry<String, Object>> valueSet = jsonObject.getMap().entrySet();

                for (Map.Entry<String, Object> entry : valueSet) {
                    String key = entry.getKey();
                    Object value = entry.getValue();

                    // Tìm chỉ số cột tương ứng với key trong dòng đầu tiên
                    int columnIndex = findColumnIndex(rowDataTemplate, key);

                    if (columnIndex != -1) {
                        // style theo template
                        CellStyle style = rowDataTemplate.getCell(columnIndex).getCellStyle();
                        // Ghi giá trị vào ô tương ứng
                        cell = row.createCell(columnIndex);
                        cell.setCellStyle(style);
                        String holder = rowDataTemplate.getCell(columnIndex).getStringCellValue();
                        if (holder.contains("<#type.NUMBER#>")) {
                            double temp = Double.parseDouble(value.toString());
                            cell.setCellValue(temp);
                        } else {
                            cell.setCellValue(value.toString());
                        }
                    }
                }
                for (int j = 0; j <= rowDataTemplate.getPhysicalNumberOfCells(); j++) {
                    cell = rowDataTemplate.getCell(j);
                    if (cell != null) {
                        String cellValue = cell.getStringCellValue();
                        if (cellValue.contains("+") || cellValue.contains("-") || cellValue.contains("*") || cellValue.contains("/")) {
                            Cell c = row.createCell(j);
                            calculatorColumn(c, jsonObject, cellValue);
                        }
                    }
                }
                for (int j = 0; j < rowDataTemplate.getLastCellNum(); j++) {
                    cell = row.getCell(j);
                    if (cell == null) {
                        cell = row.createCell(j);
                        cell.setCellStyle(valueCellStyle);
                    }
                }
            }
            sheet.removeRow(rowDataTemplate);
            sheet.shiftRows(firstRowData + 1, sheet.getLastRowNum(), -1);

            exportTempFile(sheet);

            // tính tong
            fillCellTotalValue(workbook, sheet, firstRowData, firstRowData + jsonArray.size());


        } else {
            int firstRowData = findStartRowIndex(sheet);
            Row rowDataTemplate = sheet.getRow(firstRowData);

            sheet.removeRow(rowDataTemplate);
            sheet.shiftRows(firstRowData + 1, sheet.getLastRowNum(), -1);
        }
    }


    private static void exportTempFile(XSSFSheet targetSheet) throws Exception {
        FileOutputStream fOut = new FileOutputStream("./temp.xlsx");
        targetSheet.getWorkbook().write(fOut);
        fOut.close();
    }

    private static void calculatorColumn(Cell cell, JsonObject object, String template) {
        String[] cellValues = template.split("/");
        String start = cellValues[0].replaceAll("<#table\\.(\\w+)#>", "$1");
        BigDecimal value = new BigDecimal(object.getString(start));
        for (int i = 1; i < cellValues.length; i = i + 2) {
            String key = cellValues[i + 1].replaceAll("<#table\\.(\\w+)#>", "$1");
            String keyValue = object.getString(key);
            String cal = cellValues[i];
            BigDecimal variable = new BigDecimal(keyValue);
            switch (cal) {
                case "+":
                    value = value.add(variable);
                    break;
                case "-":
                    value = value.subtract(variable);
                    break;
                case "*":
                    value = value.multiply(variable);
                    break;
                case "/":
                    value = value.divide(value);
                    break;
            }
        }
        cell.setCellValue(String.valueOf(value));
    }

    private static int findColumnIndex(Row firstRow, String key) {
        for (Cell cell : firstRow) {
            String tag = cell.getStringCellValue();
//            tag = tag.replaceAll("<#table\\.(\\w+)#>", "$1");
//            System.out.println("tag: " + tag);
            if (tag.contains(key + "#")) {
                return cell.getColumnIndex();
            }
        }
        return -1;
    }

    private static CellStyle createValueCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        setAllBorder(cellStyle);

        Font font = workbook.createFont();
        font.setFontName("Times New Roman");
        font.setFontHeightInPoints((short) 12);
        cellStyle.setFont(font);

        return cellStyle;
    }

    private static void setAllBorder(CellStyle cellStyle) {
        cellStyle.setBorderBottom(DEFAULT_BORDER_STYLE);
        cellStyle.setBorderLeft(DEFAULT_BORDER_STYLE);
        cellStyle.setBorderRight(DEFAULT_BORDER_STYLE);
        cellStyle.setBorderTop(DEFAULT_BORDER_STYLE);
    }

    public static Row getOrCreateRow(Sheet sheet, int index) {
        Row row = sheet.getRow(index);

        if (Objects.isNull(row)) {
            row = sheet.createRow(index);
        }
        return row;
    }

    public static Row createRow(Sheet sheet, int index) {
        return sheet.createRow(index);
    }

    public static Cell getOrCreateCell(Row row, Integer columnnIndex) {
        Cell cell = row.getCell(columnnIndex);

        if (Objects.isNull(cell)) {
            cell = row.createCell(columnnIndex);
        }
        return cell;
    }

    //    public static SXSSFWorkbook generateExcel(JsonArray jsonArray) {
//        if (Objects.isNull(jsonArray)) {
//            return null;
//        }
//
//        if (jsonArray.isEmpty()) {
//            throw ExportFileException.of("Không tìm thấy dữ liệu export!");
//        }
//
//        AtomicInteger atomicInteger = new AtomicInteger();
//
//        Map<String, Integer> columnNameMap
//                = jsonArray.getJsonObject(0).stream()
//                        .parallel()
//                        .collect(Collectors.toMap(Map.Entry::getKey, rs -> atomicInteger.getAndIncrement()));
//
//        SXSSFWorkbook workbook = new SXSSFWorkbook();
//
//        SXSSFSheet sheet = workbook.createSheet();
//        sheet.trackColumnsForAutoSizing(columnNameMap.values());
//
//        int headerIndexRow = 0;
//
//        Row row = sheet.createRow(headerIndexRow);
//        CellStyle headerCellType = ExcelUtils.createHeaderCellStyle(workbook);
//        Set<Map.Entry<String, Integer>> columnNameSet = columnNameMap.entrySet();
//        for (Map.Entry<String, Integer> entry : columnNameSet) {
//            Cell cell = row.createCell(entry.getValue());
//            cell.setCellStyle(headerCellType);
//            cell.setCellValue(entry.getKey());
//        }
//
////        ExcelUtils.fillData(jsonArray, columnNameMap, workbook, sheet, headerIndexRow);
//        long start = System.currentTimeMillis();
//        for (Entry<String, Integer> entry : columnNameSet) {
//            sheet.autoSizeColumn(entry.getValue());
//        }
//        Main.LOGGER.info("Autosize column take " + (System.currentTimeMillis() - start) + " ms.");
//
//        setFitWidthPage(sheet);
//
//        return workbook;
//    }
    private static CellStyle createHeaderCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        setAllBorder(cellStyle);

        Font boldFont = createBoldFont(workbook);
        cellStyle.setFont(boldFont);
        return cellStyle;
    }

    private static Font createBoldFont(Workbook workbook) {
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        return boldFont;
    }

    private static int findStartRowIndex(XSSFSheet sheet) {
//        int startTableIndex = sheet.getFirstRowNum();
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().contains("<#table.")) {
                    return row.getRowNum(); // Vị trí bắt đầu ghi dữ liệu
                }
            }
        }
        return -1; // Không tìm thấy tag
    }

    public static XSSFWorkbook formatTemplate(XSSFWorkbook templateWorkbook, JsonObject data) {
        int totalsheet = templateWorkbook.getNumberOfSheets();
        int totalRow = templateWorkbook.getSheetAt(0).getLastRowNum();

        for (int sheetIndex = 0; sheetIndex < templateWorkbook.getNumberOfSheets(); sheetIndex++) {
            Sheet sheet = templateWorkbook.getSheetAt(sheetIndex);
            for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row != null) {
                    for (int cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {
                        Cell cell = row.getCell(cellIndex);
                        if (cell != null) {
                            String cellValue;
                            try {
                                cellValue = cell.getStringCellValue();
                            } catch (Exception e) {
                                continue;
                            }
                            // Kiểm tra xem cell có chứa tag không
                            if (cellValue.contains("<#") && cellValue.contains("#>")) {
                                // Duyệt qua tất cả các key trong JsonObject
                                Set<String> setData = data.fieldNames();
                                for (String key : setData) {
                                    String tag = "<#" + key + "#>";
                                    // Nếu cell chứa tag, thì thay thế tất cả các tag bằng giá trị tương ứng
                                    while (cellValue.contains(tag)) {
                                        String value = data.getString(key);
                                        cellValue = cellValue.replaceFirst(tag, value);
                                    }
                                }
                                // Đặt giá trị đã thay thế vào cell
                                cell.setCellValue(cellValue);
                            }
                        }
                    }
                }
            }
        }
        return templateWorkbook;
    }

    public static void fillCellTotalValue(Workbook workbook, XSSFSheet sheet, int startRow, int endRow) {
        Row row = sheet.getRow(endRow);
        if (row != null) {
            Cell cell = null;
            for (int i = 0; i < row.getPhysicalNumberOfCells(); i++) {
                int temp = row.getPhysicalNumberOfCells();
                cell = row.getCell(i);
                CellStyle valueCellStyle = cell.getCellStyle();
                String tag = cell.getStringCellValue();
                if (tag != null) {
                    tag = tag.replaceAll("<#table\\.(\\w+)#>", "$1");
                }
                if (tag.equals("TOTAL")) {
                    BigDecimal total = new BigDecimal("0.0");
                    int columnIndex = cell.getColumnIndex();
                    for (int j = startRow; j <= endRow - 1; j++) {
                        Cell c = sheet.getRow(j).getCell(columnIndex);
                        String Cellname = (new CellReference(c)).formatAsString();
//                        try {
//                            double va  = c.getNumericCellValue();
//                        } catch (Exception e) {
//                            Main.LOGGER.error("getNumericCellValue Error");
//                            Main.LOGGER.error(c.getStringCellValue());
//                        }

                        // Check format cell
                        double doubleValue = 0;
                        if (c.getCellStyle().getDataFormat() == 3) {
                            try {
                                doubleValue = c.getNumericCellValue();
                            } catch (Exception e) {
                                try {
                                    doubleValue = Double.parseDouble(c.getStringCellValue());
                                } catch (Exception e1) {
//                                    Main.LOGGER.error("Value: " + c.getStringCellValue() + " is not number - catch (Calc total function - Double.parseDouble)");
                                }
//                                Main.LOGGER.error("Value: " + c.getStringCellValue() + " is not number - catch (Calc total function)");
                            }
                        } else {
//                            Main.LOGGER.error("Value: " + c.getStringCellValue() + " is not number - else (Calc total function)");
                        }

                        DecimalFormat decimalFormat = new DecimalFormat("#.####");
                        BigDecimal value = new BigDecimal(doubleValue);
                        total = total.add(value);
                    }
                    if (total.remainder(BigDecimal.ONE).compareTo(BigDecimal.ZERO) == 0) {
                        row.getCell(columnIndex).setCellValue(total.longValue());
                    } else {
                        row.getCell(columnIndex).setCellValue(total.doubleValue());
                    }
                }
            }
        }
    }
}

