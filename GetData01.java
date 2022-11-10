package Apache_Poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.print.DocFlavor.STRING;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.*;

public class GetData01 
{
	public static void main(String[] args) throws IOException 
	{
	String file ="D:\\selenium\\Apache Poi\\src\\Data\\sample sheet.xlsx";
	 FileInputStream fi  = new FileInputStream(file);
	 Workbook wb = new XSSFWorkbook(fi);
	 Sheet sh=wb.getSheetAt(1);
	 int row=sh.getLastRowNum();
	 int cell=sh.getRow(1).getLastCellNum();
	 for(int i=1;i<=row;i++) 
	 {
		 Row row1=sh.getRow(i);
		    for(int j=0;j<cell;j++) 
		    {
			   Cell cell1=row1.getCell(j);
			   switch(cell1.getCellType()) 
			   {
			   case  STRING:cell1.getStringCellValue();
			   System.out.print(cell1+ " ");
			   break;
			   case NUMERIC: cell1.getNumericCellValue();
			   System.out.print(cell1+ " ");
			   break;
			   case BOOLEAN: cell1.getBooleanCellValue();
			   System.out.print(cell1+ " ");
			   break;
			   }
			  
		 }
		  
	 }
	 wb.close();
	 }
}
