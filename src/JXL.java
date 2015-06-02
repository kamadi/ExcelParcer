import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

/**
 * Created by Tolyas on 2014-11-06.
 */
public class JXL {

    public static void read() throws IOException {
        File inputWorkbook = new File("C:/Users/Madiyar/Downloads/прайс лист алматы 27.03.15 новинки АВ и Хортекс1.xls");
        Workbook w;

        try {
            w = Workbook.getWorkbook(inputWorkbook);
            // Get the first sheet
            Sheet sheet = w.getSheet(0);
            // Loop over first 10 column and lines
            File file = new File("C:/Users/Madiyar/Downloads/divide.txt");
            BufferedWriter output = new BufferedWriter(new FileWriter(file));


//            String sql ="INSERT INTO `v_13380_shops`.`barcodes` (`lvl1`, `lvl2`, `lvl3`, `lvl4`, `lvl5`, `name`, `code`, `man`) VALUES ";
            String sql="";
            boolean start = true;
            int i = 0;
            String temp;
            for (int j = 11; j < 14; j++) {
                Cell name = sheet.getCell(3, j);
                Cell cell = sheet.getCell(4,j);
                Cell barcode = sheet.getCell(2, j);
                if (name.getContents().length() > 0 && barcode.getContents().length() > 0) {
                    if (start) {
                        sql = "INSERT INTO `product` (`id`,`name`, `barcode`) VALUES ";
                        start = false;
                    }
                    sql += "(null,";
                    String removed = removeSpace(name.getContents());
                    removed = name.getContents() +" "+cell.getContents();
                    String barcodeChanged = removeNumberSpace(barcode.getContents());

                    System.out.println(j+" "+removed + " " +barcodeChanged );

                    temp = removed.replaceAll("'", "''").replaceAll("\b","\\");
                    sql += "'" + temp + "',";
                    sql += "'" + barcodeChanged + "'";

                    if (i == 150) {
                        i = 0;
                        start = true;
                        sql += ");";
                    } else {
                        sql += "),";
                    }
//                System.out.println(sql);
                    output.write(sql);
                    output.write("\r\n");
                    i++;

                    sql = "";
                }
            }

            output.close();
        } catch (BiffException e) {
//            e.printStackTrace();
        }


    }

    public  static  String removeSpace(String text){
        int position =0;

        for (int i = 0;i<text.length();i++) {
            char ch = text.charAt(i);
            if(ch == '"' || Character.isLetter(ch)){
                position = i;
                break;
            }
        }
        return text.substring(position,text.length());
    }
    public  static  String removeNumberSpace(String text){
        int position =0;
        String result = "";
        for (int i = 0;i<text.length();i++) {
            char ch = text.charAt(i);
            if(Character.isDigit(ch)){
                result+=ch;
            }
        }
        return result;
    }
}
