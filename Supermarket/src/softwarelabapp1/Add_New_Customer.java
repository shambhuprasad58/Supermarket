/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package softwarelabapp1;

import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.PrintWriter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author SONY
 */
public class Add_New_Customer {
    public String add_customer(String First_Name, String Last_Name, String Address, String D_L_No)
    //public static void main(String args[]) throws FileNotFoundException, IOException
        {
//        String First_Name = "Shambhu";
  //      String Last_Name = "Prasad";
    //    String Address = "KGP";
      //  String D_L_No = "12345";
        
        //FileOutputStream fileOutputStream = new FileOutputStream("datalist.xlsx",true);
        //POIFSFileSystem fsFileSystem = new POIFSFileSystem(fileInputStream);
    try{
        FileInputStream myInput = new FileInputStream("datasheet.xlsx");									//Creating file input stream from desired excel file
        XSSFWorkbook wb = new XSSFWorkbook(myInput);														//Creating a temporary copy of workbook from the input stream
        XSSFSheet sh1 = wb.getSheetAt(0);																	//Using getSheetAt() function to access sheet number 1 , indexing starts from 0
        int new_CN = sh1.getLastRowNum();
        System.out.println(new_CN);
        new_CN+=1;
        int new_row = new_CN;
        // Create file
        //FileWriter fstream = new FileWriter("C:\\Lab Assignment1\\Customer_Information.txt",true);
        FileWriter fstream = new FileWriter("Customer_Information.txt",true);								//Opening a text file in Append mode
        BufferedWriter out = new BufferedWriter(fstream);													//Creating buffer to write into the text file
        out.write(new_CN+"\t"+First_Name+"\t"+Last_Name+"\t"+Address+"\t"+D_L_No);
        out.newLine();
        out.close();
        XSSFRow r1= sh1.createRow(new_row);																	//Before writing into any new box of the excel sheet which is not written before we need to create its row and its cell in the temporary workbook.
        XSSFCell c1 = r1.createCell(0);
        XSSFCell c2 = r1.createCell(1);
        c1.setCellValue(new_CN);																			//Writing into a cell
        c2.setCellValue(0.0);
        char end_column_name = 'F';
        XSSFSheet sh2 = wb.getSheetAt(1);
        set_new_row(sh2, 31, new_row, end_column_name);
        end_column_name = 'D';
        sh2 = wb.getSheetAt(2);
        set_new_row(sh2, 29, new_row, end_column_name);
        end_column_name = 'F';
        sh2 = wb.getSheetAt(3);
        set_new_row(sh2, 31, new_row, end_column_name);
        end_column_name = 'E';
        sh2 = wb.getSheetAt(4);
        set_new_row(sh2, 30, new_row, end_column_name);
        end_column_name = 'F';
        sh2 = wb.getSheetAt(5);
        set_new_row(sh2, 31, new_row, end_column_name);
        end_column_name = 'E';
        sh2 = wb.getSheetAt(6);
        set_new_row(sh2, 30, new_row, end_column_name);
        end_column_name = 'F';
        sh2 = wb.getSheetAt(7);
        set_new_row(sh2, 31, new_row, end_column_name);
        end_column_name = 'F';
        sh2 = wb.getSheetAt(8);
        set_new_row(sh2, 31, new_row, end_column_name);
        end_column_name = 'E';
        sh2 = wb.getSheetAt(9);
        set_new_row(sh2, 30, new_row, end_column_name);
        end_column_name = 'F';
        sh2 = wb.getSheetAt(10);
        set_new_row(sh2, 31, new_row, end_column_name);
        end_column_name = 'E';
        sh2 = wb.getSheetAt(11);
        set_new_row(sh2, 30, new_row, end_column_name);
        end_column_name = 'F';
        sh2 = wb.getSheetAt(12);
        set_new_row(sh2, 31, new_row, end_column_name);
        FileOutputStream fileOut = new FileOutputStream("datasheet.xlsx");				//This is most important. We need to write back the temporary workbook to our excel file or the chenges wont be reflected.
        wb.write(fileOut);																//The excel file should not be opened in the system at the time of running the code or it will give runtime error
        fileOut.close();
        String CN_no = Integer.toString(new_CN);
        int len = CN_no.length();
        for(int i=0;i<(6-len);i++)
        {
            CN_no = "0"+CN_no;
        }
    return CN_no;
  }catch (Exception e){//Catch exception if any
        System.err.println("99 addnewcustomer" + e.getMessage());
        return null;
  }
}
    private static void set_new_row(XSSFSheet sh, int days, int row, char end_column_name)
    {
        XSSFRow ROW = sh.createRow(row);
        for(int i=0; i<=days; i++)
        {
            XSSFCell cell = ROW.createCell(i);
            cell.setCellValue(0.0);
        }
        row++;
        XSSFCell cell = ROW.getCell(0);
        cell.setCellFormula("SUM(B"+row+":A"+end_column_name+row+")");
        System.out.println(cell.getNumericCellValue());
    }
}

