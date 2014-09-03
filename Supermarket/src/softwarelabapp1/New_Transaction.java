/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package softwarelabapp1;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author SONY
 */
public class New_Transaction {
    public boolean new_transaction(int CN, double Amount)
    //public static void main(String args[])
    {
        try{
        //int CN = 17;
        //double Amount = 35.0;
        double temp;
        FileInputStream myInput = new FileInputStream("datasheet.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(myInput);
        XSSFSheet sh1 = wb.getSheetAt(0);
        XSSFRow r1 = sh1.getRow(CN);
        XSSFCell c1 = r1.getCell(1);
        temp = c1.getNumericCellValue();
        System.out.println(temp);
        temp = temp + Amount;
        c1.setCellValue(temp);
        System.out.println(temp);
        r1 = sh1.getRow(0);
        c1 = r1.getCell(1);
        temp = c1.getNumericCellValue();
        System.out.println(temp);
        temp = temp + Amount;
        c1.setCellValue(temp);
        System.out.println(temp);
        DateFormat dateFormat = new SimpleDateFormat("dd");
        Date date = new Date();
        int today_date = Integer.parseInt(dateFormat.format(date));
        dateFormat = new SimpleDateFormat("MM");
        int month = Integer.parseInt(dateFormat.format(date));
        System.out.println(today_date+"     "+month);
        XSSFSheet sh2 = wb.getSheetAt(month);
        r1 = sh2.getRow(CN);
        c1 = r1.getCell(today_date);
        temp = c1.getNumericCellValue();
        temp+=Amount;
        c1.setCellValue(temp);
        c1 = r1.getCell(0);
        temp = c1.getNumericCellValue();
        temp+=Amount;
        c1.setCellValue(temp);
        r1 = sh2.getRow(0);
        c1 = r1.getCell(1);
        temp = c1.getNumericCellValue();
        temp+=Amount;
        c1.setCellValue(temp);
        FileOutputStream fileOut = new FileOutputStream("datasheet.xlsx");
        wb.write(fileOut);
        fileOut.close();
        return true;
        }catch(Exception e){//Catch exception if any
        System.err.println("68 newtransaction " + e.getMessage());
        return false;
        }
    }

}
