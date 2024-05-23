package org.example;

import io.vertx.core.json.JsonObject;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Test {
    public static void main(String[] args) throws Exception {
        File templateFile = new File("template.xlsx");
        if (!templateFile.exists()) {
            throw new Exception("Template file not found");
        }




        // Convert to POI
        XSSFWorkbook wb = new XSSFWorkbook(templateFile);

        XSSFSheet sheet = wb.getSheetAt(0);

        for (int i = 0; i < 100; i++) {
            wb.createSheet("Sheet " + i);
            System.out.println("Sheet " + i);

            XSSFSheet sheetI = wb.getSheet("Sheet " + i);

            for (int j = 0; j < 200; j++) {
                XSSFRow row = sheetI.createRow(j);
                row.createCell(0).setCellValue(j);

                sheetI.shiftRows(0, 0, 20000, true, true);
            }
        }








//
//        CellCopyPolicy cellCopyPolicy = new CellCopyPolicy();
//
////        sheet.copyRows(0, 2, 10, cellCopyPolicy);
//        int lastRow = sheet.getLastRowNum();
//        sheet.shiftRows(1, lastRow, 1, true, true);
//
        FileOutputStream fOut = new FileOutputStream("./aaaaaaaaaa.xlsx");
        wb.write(fOut);
        fOut.close();

        CellAddress cellAddress = new CellAddress(0, 0);

        System.out.println(cellAddress.toString());

        System.out.println("12312");
    }
}
