package org.example;


import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;

public class Test {
    public static void main(String[] args) throws Exception {
        //  KH_15_20240624170201.docx
        //  Mau 33C-THQ.docx

        File templateFile = new File("KH_15_20240624170201.docx");
        if (!templateFile.exists()) {
            throw new Exception("Template file not found");
        }
        XWPFDocument doc = new XWPFDocument(OPCPackage.open(templateFile));

        XWPFComment[] a = doc.getComments();

        removeAllComment(doc);
//        Workbook workbook = new Workbook();
//        workbook.loadFromFile("KH_02.xlsx");
//
//        //Get the second worksheet
//        Worksheet worksheet = workbook.getWorksheets().get(0);
//
//        //Save as PDF document
//        worksheet.saveToPdf("KH_02.pdf");

        FileOutputStream fOut = new FileOutputStream("./AAAAAresult.docx");
        doc.write(fOut);
        fOut.close();
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