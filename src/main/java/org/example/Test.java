package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;

public class Test {
    public static void main(String[] args) throws Exception {
        File templateFile = new File("template.xlsx");
        if (!templateFile.exists()) {
            throw new Exception("Template file not found");
        }

//        CreationHelper factory = workbook.getCreationHelper();


        // Convert to POI
        XSSFWorkbook wb = new XSSFWorkbook(templateFile);

        XSSFSheet sheet = wb.getSheetAt(0);

        XSSFRow sheetRow = sheet.getRow(6);
        XSSFCell cell = sheetRow.getCell(1);

//        XSSFComment comment = new XSSFComment("asdawdawdawd", CTComment.);

        CreationHelper factory = wb.getCreationHelper();
        //get an existing cell or create it otherwise:

        ClientAnchor anchor = factory.createClientAnchor();

        Drawing<XSSFShape> drawing = sheet.createDrawingPatriarch();
        Comment comment = drawing.createCellComment(anchor);
        //set the comment text and author
        comment.setString(factory.createRichTextString("createRichTextString"));
        comment.setAuthor("setAuthor");

        cell.setCellComment(comment);

//        JsonObject jsonObjectComment = new JsonObject();
//
//        HashMap<String, MergeCellList> mergeCellLists = new HashMap<>();
//        String key = "";
//
//        for (int i = 0; i < sheet.getLastRowNum(); i++) {
//            XSSFRow row = sheet.getRow(i);
//            if (row == null) {
//                continue;
//            }
//
//            for (int j = 0; j < row.getLastCellNum(); j++) {
//                XSSFCell cell = row.getCell(j);
//                if (cell == null) {
//                    continue;
//                }
//
//                XSSFComment comment = cell.getCellComment();
//                if (comment == null) {
//                    continue;
//                }
//
//                String commentValue = String.valueOf(comment.getString());
//                boolean hasKey = mergeCellLists.containsKey(commentValue);
//                if (!hasKey) {
//                    key = commentValue;
//                    MergeCellList mergeCellList = new MergeCellList(commentValue);
//                    mergeCellList.addCell(new CellAddress(cell));
//                    mergeCellLists.put(commentValue, mergeCellList);
//                } else {
//                    MergeCellList mergeCellList = mergeCellLists.get(commentValue);
//                    mergeCellList.addCell(new CellAddress(cell));
//                    mergeCellLists.put(commentValue, mergeCellList);
//                }
//
//                String valueCell = cell.getStringCellValue();
//                if (valueCell != null && commentValue.contains("(empty)")) {
//                    cell.setCellValue("");
//                }
//                cell.removeCellComment();
//            }
//        }
//
//        CellRangeAddress cellAddresses = mergeCellLists.get(key).getCellRangeAddress();
//
//        sheet.addMergedRegion(cellAddresses);

        FileOutputStream fOut = new FileOutputStream("./temp.xlsx");
        sheet.getWorkbook().write(fOut);
        fOut.close();

        System.out.println("Done test!!");
    }
}
