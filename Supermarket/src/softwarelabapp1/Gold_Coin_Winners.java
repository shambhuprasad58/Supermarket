/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package softwarelabapp1;

import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.PrintWriter;
import java.text.DateFormat;
import java.text.DateFormatSymbols;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author SONY
 */
public class Gold_Coin_Winners {
    public void winners(){
    try{
        double [][] highest_purchase_winner_list = new double[10][2];
        double total_sale = 0.0;
        for(int i=0;i<10;i++)
        {
            highest_purchase_winner_list[i][0]=0;
            highest_purchase_winner_list[i][1]=0;
        }
        DateFormat dateFormat = new SimpleDateFormat("yyyy");
        Date year = new Date();
        int this_year = Integer.parseInt(dateFormat.format(year));
        //FileWriter fstream = new FileWriter("C:\\Lab Assignment1\\Gold Coin Winners "+this_year+".txt");
        FileWriter fstream = new FileWriter("Gold Coin Winners "+this_year+".txt");
        PrintWriter out = new PrintWriter(fstream);
        out.write("GOLD COIN WINNERS");
        out.println();
        out.write("CN      \t PURCHASE IN "+this_year);
        out.println();
        FileInputStream myInput = new FileInputStream("datasheet.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(myInput);
        XSSFSheet sh1 = wb.getSheetAt(0);
        Iterator row_iter = sh1.rowIterator();
        XSSFRow r1;
        XSSFCell c1,c2;
        row_iter.next();
        while(row_iter.hasNext())
        {
        r1 =(XSSFRow) row_iter.next();
        c1 = r1.getCell(1);
        System.out.println(c1.getNumericCellValue());
        total_sale+=c1.getNumericCellValue();
        if(c1.getNumericCellValue()>=highest_purchase_winner_list[0][1])
        {
             highest_purchase_winner_list[0][1]=c1.getNumericCellValue();
             c2 = r1.getCell(0);
             highest_purchase_winner_list[0][0]=c2.getNumericCellValue();
             correct_position(highest_purchase_winner_list);
        }
        if(c1.getNumericCellValue()>9999.99)
        {
            c2 = r1.getCell(0);
            out.write(c2.getNumericCellValue()+"\t");
            out.write(""+c1.getNumericCellValue());
            out.println();
        }
        }
        out.close();
        //fstream = new FileWriter("C:\\Lab Assignment1\\Highest Purchase Winner "+this_year+".txt");
        fstream = new FileWriter("Highest Purchase Winner "+this_year+".txt");
        out = new PrintWriter(fstream);
        out.write("HIGHEST PURCHASE WINNERS");
        out.println();
        out.write("CN      \t PURCHASE IN "+this_year);
        out.println();
        for(int i=9;i>=0;i--)
        {
            out.write(highest_purchase_winner_list[i][0]+"\t"+highest_purchase_winner_list[i][1]);
            out.println();
        }
        out.close();
        System.out.println("85 gold coin correct");
        //fstream = new FileWriter("C:\\Lab Assignment1\\Total Sales Position "+this_year+".txt");
        fstream = new FileWriter("Total Sales Position "+this_year+".txt");
        out = new PrintWriter(fstream);
        r1 = sh1.getRow(0);
        c1 = r1.getCell(1);
        out.write("Total sale of year "+this_year+" is of Rs"+c1.getNumericCellValue());
        System.out.println("91 gold coin correct");
        out.println();
        out.write("Monthly Sales Of "+this_year);
        out.println();
        for(int i=1;i<13;i++)
        {
            sh1 = wb.getSheetAt(i);
            r1 = sh1.getRow(0);
            c1 = r1.getCell(1);
            out.write(new DateFormatSymbols().getMonths()[i-1]+"\t"+c1.getNumericCellValue());
            out.println();
        }
        out.close();
    }catch(Exception e){//Catch exception if any
        System.err.println("102 goldcoinwinner" + e.getMessage());
        }
    }
    private static void correct_position(double Arr[][]){
        double temp;
        for(int i=0;i<9;i++)
        {
            if(Arr[i][1]>=Arr[i+1][1])
            {
                temp=Arr[i][1];
                Arr[i][1]= Arr[i+1][1];
                Arr[i+1][1]=temp;
                temp=Arr[i][0];
                Arr[i][0]= Arr[i+1][0];
                Arr[i+1][0]=temp;
                System.out.println(Arr[i+1][1]);
           //     Arr[i][1]+=Arr[i+1][1]-(Arr[i+1][1]=Arr[i][1]);
             //   Arr[i][0]+=Arr[i+1][0]-(Arr[i+1][0]=Arr[i][0]);
            }
            else
                break;
        }
    }
}
