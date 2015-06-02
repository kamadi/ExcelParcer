/**
 * Created by Madiyar on 02.06.2015.
 */

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;

import java.io.*;
import java.util.Iterator;

public class ApachePOI {
    public static void read() {
        try {
            File divide = new File("C:/Users/Madiyar/Downloads/divide3.txt");
            BufferedWriter output = new BufferedWriter(new FileWriter(divide));
            String sql = "";
            boolean start = true;
            final DataFormatter df = new DataFormatter();
            FileInputStream file = new FileInputStream(new File("C:/Users/Madiyar/Downloads/прайс-1.xls"));

            //Get the workbook instance for XLS file
            HSSFWorkbook workbook = new HSSFWorkbook(file);

            //Get first sheet from the workbook
            HSSFSheet sheet = workbook.getSheetAt(0);

            //Iterate through each rows from first sheet
            Iterator<Row> rowIterator = sheet.iterator();
            int i = 1;
            int j = 1;
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                System.out.print(j + " ");
                if (j > 15) {
                    Cell barcodeCell = row.getCell(1);

                    Cell nameCell1 = row.getCell(2);
//                    Cell nameCell2 = row.getCell(4);
                    String barcode = "", name = "";
                    try {
                        barcode = removeNumberSpaceTemp(df.formatCellValue(barcodeCell));
                        name = removeSpace(nameCell1.getStringCellValue());
                    } catch (NullPointerException e) {
                        continue;
                    }

                    if (!barcode.isEmpty() && !name.isEmpty()) {
                        if (start) {
                            sql = "INSERT INTO `product` (`id`,`name`, `barcode`) VALUES ";
                            start = false;
                        }
                        sql += "(null,";
                        sql += "'" + name + "',";
                        sql += "'" + barcode + "'";

                        if (i == 150) {
                            i = 0;
                            start = true;
                            sql += ");";
                        } else {
                            sql += "),";
                        }
//                    System.out.println(sql);
                        System.out.print(name + " " + barcode);
                        output.write(sql);
                        output.write("\r\n");
                        sql = "";
                    }
                }
                i++;
                j++;
                System.out.println();
            }
            file.close();
            output.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    public static String removeSpace(String text) {
        int position = 0;

        for (int i = 0; i < text.length(); i++) {
            char ch = text.charAt(i);
            if (ch == '"' || Character.isLetter(ch)) {
                position = i;
                break;
            }
        }
        return text.substring(position, text.length()).replaceAll("'", "''").replaceAll("\b", "\\").trim();
    }

    public static String removeNumberSpace(String text) {
        int position = 0;
        String result = "";
        for (int i = 0; i < text.length(); i++) {
            char ch = text.charAt(i);
            if (Character.isDigit(ch)) {
                result += ch;
            } else {
                break;
            }
        }
        return result;
    }

    public static String removeNumberSpaceTemp(String text) {
        int position = 0;
        String result = "";
        for (int i = 0; i < text.length(); i++) {
            char ch = text.charAt(i);
            if (ch == '/') {
                if (i == text.length() - 1) {
                    break;
                }
                else {
                    position = i;
                    break;
                }
            }
        }
        if(text.isEmpty())
            return "";

        return text.substring(position + 1, text.length()).replaceAll("'", "''").replaceAll("\b", "\\").trim();
    }
}
