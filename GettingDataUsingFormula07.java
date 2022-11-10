package Apache_Poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GettingDataUsingFormula07 
{
	
    public static FileInputStream fi;
    public static Workbook wb;
    public static Sheet sh;
    public static Row row;
    public static Cell cell;
    
	public static void main(String[] args) throws IOException 
	{
		String file="D:\\selenium\\Apache Poi\\src\\Data\\ExcelFormula.xlsx";
		fi= new FileInputStream(file);
		wb = new XSSFWorkbook(fi);
		sh = wb.getSheet("Sheet1");
	   int rw=sh.getLastRowNum();
	    int cl=sh.getRow(1).getLastCellNum();
	    
	    for(int r=0;r<rw;r++) 
	    {
	      row=sh.getRow(r);
	    	 for(int c=0;c<cl;c++) {
	    		 cell=row.getCell(c);
	    		 switch(cell.getCellType()) {
	    		 case STRING:
	    			 System.out.print(cell.getStringCellValue());
	    			 break;
	    		 case NUMERIC:
	    			 System.out.print(cell.getNumericCellValue());
	    			 break;
	    		 case BOOLEAN:
	    			 System.out.print(cell.getBooleanCellValue());
	    			 break;
	    		 case FORMULA:
	    			 System.out.print(cell.getNumericCellValue());
	    			 break;
	    		 }
	    		 System.out.print(" | ");
               }
	    	 System.out.println();
	    }
		
		

	}

}
