/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package softwarelabapp1;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author SONY
 */
public class Reset_All {
    public void reset_all()
    {
try{
    Gold_Coin_Winners GCW = new Gold_Coin_Winners();
    GCW.winners();
    FileInputStream myInput = new FileInputStream("datasheet.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(myInput);
        XSSFSheet sh1 = wb.getSheetAt(0);
        Iterator row_iter = sh1.rowIterator();
        XSSFRow r1;
        XSSFCell c1;
        while(row_iter.hasNext())
        {
            r1 = (XSSFRow) row_iter.next();
            r1.getCell(1).setCellValue(0.0);
        }
        Iterator cell_iter;
        for(int i=1;i<13;i++)
        {
            sh1 = wb.getSheetAt(i);
            row_iter = sh1.rowIterator();
            while(row_iter.hasNext())
            {
                r1 = (XSSFRow) row_iter.next();
                cell_iter = r1.cellIterator();
                while(cell_iter.hasNext())
                {
                    c1 = (XSSFCell) cell_iter.next();
                    c1.setCellValue(0.0);
                }
            }
        }
        FileOutputStream fileOut = new FileOutputStream("datasheet.xlsx");
        wb.write(fileOut);
        fileOut.close();
    }catch(Exception e){
        System.out.println("52 resetall"+e);
    }
}
}
