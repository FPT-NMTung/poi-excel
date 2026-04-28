package org.example;

import converter.Converter;

import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.text.NumberFormat;
import java.util.Arrays;
import java.util.Locale;

public class MainSxssFile {
    public static void main(String[] args) throws Exception {
        System.out.println(formatField("0000.01230100000", "number2"));
    }

    private static String formatField(String value, String format) {
        String result = "";
        BigDecimal bd;
        if (format == null) {
            return value;
        }

        switch (format) {
            case "number":
                DecimalFormatSymbols symbols = DecimalFormatSymbols.getInstance();
                symbols.setGroupingSeparator(',');

                DecimalFormat formatter = new DecimalFormat("###,###.########", symbols);

                return formatter.format(new BigDecimal(value));
            case "number_char_vi":
                try {
                    bd = new BigDecimal(value);
                    result = Converter.numberToCharVi(bd.stripTrailingZeros().toPlainString()).trim();
                } catch (Exception e) {
                    result = value;
                }
                break;
            case "number_char_Vi":
                try {
                    bd = new BigDecimal(value);
                    result = Converter.numberToCharVi(bd.stripTrailingZeros().toPlainString()).trim();
                    if (result.length() >= 1) {
                        result = result.substring(0, 1).toUpperCase() + result.substring(1);
                    }
                } catch (Exception e) {
                    result = value;
                }
                break;
            case "number_char_VI":
                try {
                    bd = new BigDecimal(value);
                    result = Converter.numberToCharVi(bd.stripTrailingZeros().toPlainString()).trim().toUpperCase();
                } catch (Exception e) {
                    result = value;
                }
                break;
            case "number_char_en":
                try {
                    result = Converter.numberToCharEn(value).trim();
                } catch (Exception e) {
                    result = value;
                }
                break;
            case "number_char_En":
                try {
                    result = Converter.numberToCharEn(value).trim();
                    if (result.length() >= 1) {
                        result = result.substring(0, 1).toUpperCase() + result.substring(1);
                    }
                } catch (Exception e) {
                    result = value;
                }
                break;
            case "number_char_EN":
                try {
                    result = Converter.numberToCharEn(value).trim().toUpperCase();
                } catch (Exception e) {
                    result = value;
                }
                break;
            case "checkbox":
                switch (value) {
                    case "TICK_V":
                        result = "\uF052";
                        break;
                    case "TICK_X":
                        result = "\uF051";
                        break;
                    default:
                        result = "\uF0A3";
                        break;
                }

                break;
            default:
                result = value;
        }

        return result;
    }
}
