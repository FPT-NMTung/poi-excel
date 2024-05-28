package converter;

import java.util.ArrayList;
import java.util.List;

public class Converter {
    public static void main(String[] args) {
//        System.out.println(numberToCharVi("0"));
//        System.out.println(numberToCharVi("123"));
//        System.out.println(numberToCharVi("12345"));
//        System.out.println(numberToCharVi("123456"));
//        System.out.println(numberToCharVi("123456789"));
        System.out.println(numberToCharVi("99111222333444555666"));
    }

    public static String numberToCharVi(String strNumber) {
        // Check exist number
        if (strNumber == null || strNumber.isEmpty()) {
            return "";
        }

        StringBuilder result = new StringBuilder();

        // Split string
        int maxSplit = (int) Math.ceil((double) strNumber.length() /9);
        int intDiv = strNumber.length() %9;

        // loop process
        for (int index = 1; index <= maxSplit; index++) {
            if (index == maxSplit) {
                String subNumber = strNumber.substring(0, intDiv == 0 ? 9 : intDiv);
                StringBuilder subNumberConverted = new StringBuilder();

                int intSubDiv = subNumber.length() %3;

                if (subNumber.length() > 6) {
                    String tempResult = Converter.numberToCharViSub(subNumber.substring(0, intSubDiv == 0 ? 3 : intSubDiv), "", " triệu");
                    subNumberConverted.append(tempResult);

                    if (intSubDiv == 0) {
                        intSubDiv = 3;
                    }

                    tempResult = Converter.numberToCharViSub(subNumber.substring(intSubDiv, intSubDiv + 3), "", " nghìn");
                    subNumberConverted.append(tempResult);

                    tempResult = Converter.numberToCharViSub(subNumber.substring(intSubDiv + 3, intSubDiv + 3 + 3), "", "");
                    subNumberConverted.append(tempResult);
                }

                if (subNumber.length() > 3 && subNumber.length() <= 6) {
                    String tempResult = Converter.numberToCharViSub(subNumber.substring(0, intSubDiv == 0 ? 3 : intSubDiv), "", " nghìn");
                    subNumberConverted.append(tempResult);

                    if (intSubDiv == 0) {
                        intSubDiv = 3;
                    }

                    tempResult = Converter.numberToCharViSub(subNumber.substring(intSubDiv, intSubDiv + 3), "", "");
                    subNumberConverted.append(tempResult);
                }

                if (subNumber.length() <= 3) {
                    String tempResult = Converter.numberToCharViSub(subNumber, "", "");
                    subNumberConverted.append(tempResult);
                }

                result.insert(0, subNumberConverted.toString());
            } else {
                int startIndex;
                if (intDiv == 0) {
                    startIndex = (maxSplit - index) * 9;
                } else {
                    startIndex = 9 * (maxSplit - index) - 9 + intDiv;
                }

                String subNumber = strNumber.substring(startIndex, startIndex + 9);// (maxSplit - index)
                StringBuilder subNumberConverted = new StringBuilder();

                String tempResult = Converter.numberToCharViSub(subNumber.substring(0, 3), "", " triệu");
                subNumberConverted.append(tempResult);

                tempResult = Converter.numberToCharViSub(subNumber.substring(3, 6), "", " nghìn");
                subNumberConverted.append(tempResult);

                tempResult = Converter.numberToCharViSub(subNumber.substring(6, 9), "", "");
                subNumberConverted.append(tempResult);

                result.insert(0, " tỷ" + subNumberConverted.toString());
            }
        }

        return result.toString();
    }

    private static String numberToCharViSub(String strNumber, String prefix, String suffix) {
        StringBuilder resultString = new StringBuilder();

        int length = strNumber.length();
        String digit1 = strNumber.length() > 0 ? strNumber.substring(0, 1) : null;
        String digit2 = strNumber.length() > 1 ? strNumber.substring(1, 2) : null;
        String digit3 = strNumber.length() > 2 ? strNumber.substring(2, 3) : null;

        if (strNumber.equals("000")) {
            return resultString.toString();
        }

        resultString.append(prefix);

        switch (digit1) {
            case "0":
                if (length == 3 || length == 1) {
                    resultString.append(" không");
                }
                break;
            case "1":
                if (length == 2) {
                    resultString.append(" mười");
                } else {
                    resultString.append(" một");
                }
                break;
            case "2":
                resultString.append(" hai");
                break;
            case "3":
                resultString.append(" ba");
                break;
            case "4":
                resultString.append(" bốn");
                break;
            case "5":
                resultString.append(" năm");
                break;
            case "6":
                resultString.append(" sáu");
                break;
            case "7":
                resultString.append(" bảy");
                break;
            case "8":
                resultString.append(" tám");
                break;
            case "9":
                resultString.append(" chín");
                break;
        }

        if (length == 3) {
            resultString.append(" trăm");
        }

        if (length == 2 && !digit1.equals("0") && !digit1.equals("1")) {
            resultString.append(" mươi");
        }

        if (length >= 2) {
            switch (digit2) {
                case "1":
                    if (length == 3) {
                        resultString.append(" mười");
                    } else if (length == 2 && !digit1.equals("1")) {
                        resultString.append(" mốt");
                    } else if (length == 2 && !digit1.equals("0")) {
                        resultString.append(" một");
                    }
                    break;
                case "2":
                    resultString.append(" hai");
                    break;
                case "3":
                    resultString.append(" ba");
                    break;
                case "4":
                    if (length == 2 && digit1 != "0" && digit1 != "1") {
                        resultString.append(" tư");
                    } else {
                        resultString.append(" bốn");
                    }
                    break;
                case "5":
                    if (length == 2) {
                        resultString.append(" lăm");
                    } else {
                        resultString.append(" năm");
                    }
                    break;
                case "6":
                    resultString.append(" sáu");
                    break;
                case "7":
                    resultString.append(" bảy");
                    break;
                case "8":
                    resultString.append(" tám");
                    break;
                case "9":
                    resultString.append(" chín");
                    break;
            }

            if (length == 3) {
                if (!digit2.equals("0") && !digit2.equals("1")) {
                    resultString.append(" mươi");
                }
                if (digit2.equals("0") && !digit3.equals("0")) {
                    resultString.append(" linh");
                }
            }
        }

        if (length == 3) {
            switch (digit3) {
                case "1":
                    if (!digit2.equals("0") && !digit2.equals("1")) {
                        resultString.append(" mốt");
                    } else {
                        resultString.append(" một");
                    }
                    break;
                case "2":
                    resultString.append(" hai");
                    break;
                case "3":
                    resultString.append(" ba");
                    break;
                case "4":
                    if (length == 2 && !digit2.equals("0") && !digit2.equals("1")) {
                        resultString.append(" tư");
                    } else {
                        resultString.append(" bốn");
                    }
                    break;
                case "5":
                    if (!digit2.equals("0")) {
                        resultString.append(" lăm");
                    } else {
                        resultString.append(" năm");
                    }
                    break;
                case "6":
                    resultString.append(" sáu");
                    break;
                case "7":
                    resultString.append(" bảy");
                    break;
                case "8":
                    resultString.append(" tám");
                    break;
                case "9":
                    resultString.append(" chín");
                    break;
            }
        }

        resultString.append(suffix);

        return resultString.toString();
    }

    public static String numberToCharEn(String number) {
        ConverterEn.DefaultProcessor processor = new ConverterEn.DefaultProcessor();
        return processor.getName(number);
    }
}
