package Apache_Poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GetProtectData06 
{
	public static FileInputStream fi;
	public static XSSFWorkbook  wb;
	public static XSSFSheet sh;
	public static XSSFRow rw;
	public static XSSFCell cell;

	public static void main(String[] args) throws IOException 
	{
		String path="D:\\selenium\\Apache Poi\\src\\Data\\Encrypt.xlsx";
		String pwd="Pavan1234";
		fi= new FileInputStream(path);
		wb=(XSSFWorkbook) WorkbookFactory.create(fi,pwd );
		sh=wb.getSheetAt(0);
	    Iterator<Row> itr=sh.iterator();
	    while(itr.hasNext()) {
	    	rw=(XSSFRow) itr.next();
	    	Iterator clitr=rw.iterator();
	    	while(clitr.hasNext()) {
	    		cell=(XSSFCell) clitr.next();
	    		switch(cell.getCellType()) 
	    		{
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
		wb.close();
	}

}
