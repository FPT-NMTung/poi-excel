package org.example;


import com.spire.xls.CellStyle;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;

import javax.swing.text.DateFormatter;
import java.io.File;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

public class Test {
    public static void main(String[] args) throws Exception {
        //  KH_15_20240624170201.docx
        //  Mau 33C-THQ.docx

        // Read template
//        File templateFile = new File("HD_MO_TK.docx");
//        if (!templateFile.exists()) {
//            throw new Exception("Template file not found");
//        }
//
//        XWPFDocument doc = new XWPFDocument(OPCPackage.open(templateFile));

        // Get file template
        File templateFile = new File("test-format.xlsx");
//        File templateFile = new File("GD_07.xlsx");

        if (!templateFile.exists()) {
            throw new Exception("Template file not found");
        }

        long startTime;
        long beginTime = System.currentTimeMillis();

        // Convert to POI
        XSSFWorkbook wb = new XSSFWorkbook(templateFile);

        XSSFSheet sheet = wb.getSheetAt(0);
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");

        XSSFRow row = sheet.getRow(0);
        XSSFCell cell = row.getCell(0);

        XSSFCellStyle cellStyle = cell.getCellStyle();
        System.out.println(cellStyle.getDataFormatString());
        System.out.println(cellStyle.getDataFormat());

        cell.setCellValue(dateFormat.parse("31/12/2023"));

        row = sheet.getRow(1);
        cell = row.getCell(0);

        cell.setCellValue(dateFormat.parse("31/12/2023"));

//        DataFormat dataFormat = new DataFormat().;
        System.out.println(cellStyle.getDataFormat());

        FileOutputStream fOut = new FileOutputStream("./result.xlsx");
        wb.write(fOut);
        fOut.close();

//        int count = 0;
//        int target = -1;
//        for (XWPFParagraph paragraph : doc.getParagraphs()) {
//            String contentPara = paragraph.getText();
//            System.out.println(count + " " + contentPara);
//
//            if (contentPara.trim().equals("<#DELETE_LINE>")) {
//                target = count;
//            }
//
//            count ++;
//        }

//        doc.removeBodyElement(3);
//        doc.removeBodyElement(3);
//        doc.removeBodyElement(3);

//        for (int i = doc.getBodyElements().size() - 1; i >= 0; i--) {
//            IBodyElement iBodyElement = doc.getBodyElements().get(i);
//            if (iBodyElement.getElementType() == BodyElementType.PARAGRAPH) {
//                XWPFParagraph paragraph = (XWPFParagraph) iBodyElement;
//                System.out.println(i + " " + paragraph.getText());
//
//                if (paragraph.getText().trim().equals("<#DELETE_LINE>")) {
//                    doc.removeBodyElement(i);
//                }
//            }
//        }
//
//
//
//        FileOutputStream fOut = new FileOutputStream("./AAAAAresult.docx");
//        doc.write(fOut);
//        fOut.close();
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
}