package org.example;

import converter.Converter;
import converter.ConverterEn;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFComment;
import org.apache.poi.xwpf.usermodel.XWPFComments;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.io.File;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Test {
    public static void main(String[] args) throws Exception {
//        File templateFile = new File("Mau 33C-THQ.docx");
//        if (!templateFile.exists()) {
//            throw new Exception("Template file not found");
//        }
//
//        XWPFDocument doc = new XWPFDocument(OPCPackage.open(templateFile));
//
//        List<XWPFTable> table = doc.getTables();
//        String sourceNumber = "0111222333444555666777";
//
//        ConverterEn.DefaultProcessor processor = new ConverterEn.DefaultProcessor();
//        String val = processor.getName(sourceNumber);
//
//        System.out.println(val);
//
//
//        DecimalFormat formatter = new DecimalFormat("#,###.0000");
//        System.out.println(formatter.format(Double.parseDouble("123123123123.01312")));




        System.out.println("12312.00".replaceAll("[.]0+", ""));

        System.out.println("Done test!!");

        BigDecimal bd  = new BigDecimal("23.10");
        BigDecimal bd1 = new BigDecimal("0.99000000000000000000000");

        bd  = bd.stripTrailingZeros();
        bd1 = bd1.setScale(6, RoundingMode.HALF_UP).stripTrailingZeros();

        System.out.println("bd value::"+ bd.stripTrailingZeros().toPlainString()); System.out.println("bd1 value::"+ bd1);
    }
}