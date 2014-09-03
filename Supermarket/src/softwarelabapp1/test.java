/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package softwarelabapp1;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.text.DateFormat;
import java.text.DateFormatSymbols;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import org.apache.poi.hslf.model.Sheet;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author SONY
 */
public class test {
public static void main(String args[]) throws Exception {
    /*String fileName = "D:\\Studies\\2nd Year\\4th Sem\\Software\\SoftwareLab App1\\datasheet.xlsx";
    FileInputStream myInput = new FileInputStream(fileName);
    //POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);
    //XSSFWorkbook myWorkbook;
    XSSFWorkbook myWorkbook = new XSSFWorkbook(myInput);
    //HSSFWorkbook myWorkBook = new HSSFWorkbook(myFileSystem);
    XSSFSheet mySheet = myWorkbook.getSheetAt(0);
    XSSFRow r2 = mySheet.createRow(1);
    //HSSFRow r1 = mySheet.getRow(0);
    XSSFCell c1 = r2.createCell(0);
    Iterator rowIter = mySheet.rowIterator();
    Cell cell = r2.getCell(0);
    cell.setCellValue(476.876);
    //Modify the cellContents here
    // Write the output to a file
    //cell.setCellValue(cellContents);
    FileOutputStream fileOut = new FileOutputStream("D:\\Studies\\2nd Year\\4th Sem\\Software\\SoftwareLab App1\\datasheet.xlsx");
    myWorkbook.write(fileOut);
    fileOut.close();
    FileWriter outFile = new FileWriter("C:\\Lab Assignment1\\Gold Coin Winners 12.txt");
    PrintWriter out = new PrintWriter(outFile);
    double d = 132.0;
    int i = (int) d;
    out.write(i);
    out.println();
    out.close();
    /*FileOutputStream fos = new FileOutputStream("C:\\Lab Assignment1\\Gold Coin Winners 12.txt");
    OutputStreamWriter out = new OutputStreamWriter(fos);
    out.write("new2\n");
    out.close();*/
    DateFormat dateFormat = new SimpleDateFormat("yyyy");
        Date date = new Date();
        int today_date = Integer.parseInt(dateFormat.format(date));
        dateFormat = new SimpleDateFormat("MM");
        int month = Integer.parseInt(dateFormat.format(date));
        Calendar cal = Calendar.getInstance();
        String len = "Total sale of year "+today_date+" is of Rs";
        System.out.println(len+len.length());
        System.out.println(today_date+"  "+new DateFormatSymbols().getMonths()[11]+"   "+month);
    String str = "0001";
    int a = Integer.parseInt(str);
    System.out.println(a);
}
}
