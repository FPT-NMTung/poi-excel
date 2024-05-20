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

public class Test {
    public static void main(String[] args) throws Exception {
        File templateFile = new File("template.xlsx");
        if (!templateFile.exists()) {
            throw new Exception("Template file not found");
        }

        // Convert to POI
        XSSFWorkbook wb = new XSSFWorkbook(templateFile);

        XSSFSheet sheet = wb.getSheetAt(0);

//        XSSFRow row = sheet.getRow(2);
//        XSSFCell cell = row.getCell(4);

        CellAddress ca = new CellAddress("E10");

        System.out.println(ca.getRow());
        System.out.println(ca.getColumn());

        CellCopyPolicy cellCopyPolicy = new CellCopyPolicy();

//        sheet.copyRows(1, 3, 10, cellCopyPolicy);

//        FileOutputStream fOut = new FileOutputStream("./aaaaaaaaaa.xlsx");
//        wb.write(fOut);
//        fOut.close();

        System.out.println("123");
    }
}
