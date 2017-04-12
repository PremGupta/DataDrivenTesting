package Exceldriven;

/*import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDrivenTestCase {

	public static FileInputStream fis;
	public static XSSFWorkbook wb;
	public static XSSFSheet sheet;
	public static XSSFRow row;
	public static XSSFRow row1;
	public static XSSFCell cell;
	
	public static void main(String[] args) throws IOException 
	{

		 String value1=getCelldata(1,2);
		 System.out.println(value1);
		 
		 value1= setCelldata("automation ", 1,2);
		 System.out.println(value1);
		
	}    
		



	public static String getCelldata(int rownum, int col) throws IOException
	{
		
		 fis= new FileInputStream("E:\\Mee\\UserLoginDetails.xlsx");
		 wb= new XSSFWorkbook(fis);
         sheet= wb.getSheet("Sheet1");
         row= sheet.getRow(1);
         cell= row.getCell(2);
        //String value= cell.getStringCellValue();
        //System.out.println(value);
      
        String celldata1=cell.getStringCellValue();
        return celldata1;
        
        //return cell.getStringCellValue();
		
	}
	
	public static String setCelldata(String text, int rownum, int col) throws IOException
	{
		
		 fis= new FileInputStream("E:\\Mee\\UserLoginDetails.xlsx");
		 wb= new XSSFWorkbook(fis);
         sheet= wb.getSheet("Sheet1");
         row= sheet.getRow(1);
         cell= row.getCell(2);
       
      
         cell.setCellValue(text);
        String celldata1=cell.getStringCellValue();
        
        return celldata1;

	}
	
	
}*/

import java.io.FileInputStream;

import java.io.FileNotFoundException;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;

import org.apache.poi.xssf.usermodel.XSSFRow;

import org.apache.poi.xssf.usermodel.XSSFSheet;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.openqa.selenium.By;



public class ExcelDrivenTestCase 
{

public static XSSFWorkbook wb;
public static XSSFSheet sheet;
public static XSSFRow row;
public static XSSFCell cell;
public static FileInputStream fis;
public static void main(String[] args) throws IOException {


/*String value=getCelldata(2,2);
System.out.println(value);*/
String value1=getCelldata(1,2);
System.out.println(value1);
value1=setCelldata("automation",2,2);
System.out.println(value1);

}


private static String setCellvalue(int i, int j) 
{
return null;
}

//Getter() method
public static String getCelldata( int rownum,int col) throws IOException
{
fis =new FileInputStream("C:\\Users\\Prem\\Downloads\\Study\\Selenium Files\\UserLoginDetails.xlsx");
wb=new XSSFWorkbook(fis);
sheet=wb.getSheet("sheet1");
row=sheet.getRow(1);
cell=row.getCell(2);
return cell.getStringCellValue();
}


//Setter() method
public static String setCelldata(String text,int rownum,int col) throws IOException

{
fis =new FileInputStream("C:\\Users\\Prem\\Downloads\\Study\\Selenium Files\\UserLoginDetails.xlsx");
wb=new XSSFWorkbook(fis);
sheet=wb.getSheet("sheet1");
row=sheet.getRow(1);
cell=row.getCell(2);
cell.setCellValue(text);
String celldata1= cell.getStringCellValue();
return celldata1;
}

}