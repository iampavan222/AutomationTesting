package Apache_Poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GetDataUsingIterator02 {

	public static FileInputStream fi;
	public static Workbook wb;
	public static void main(String[] args) throws IOException {
		String file ="D:\\selenium\\Apache Poi\\src\\Data\\sample sheet.xlsx";
		 fi  = new FileInputStream(file);
		 wb = new XSSFWorkbook(fi);
		 Sheet sh=wb.getSheetAt(1);
		 
		 Iterator itr =sh.iterator();
		 while(itr.hasNext()) {
			 Row row=(Row) itr.next();
			 Iterator itrcell=row.cellIterator();
			while(itrcell.hasNext()) {
				Cell cell=(Cell) itrcell.next();
				switch(cell.getCellType()) 
				   {
				   case  STRING:cell.getStringCellValue();
				   System.out.println(cell+ " | ");
				   break;
				   case NUMERIC: cell.getNumericCellValue();
				   System.out.println(cell+ " | ");
				   break;
				   case BOOLEAN: cell.getBooleanCellValue();
				   System.out.println(cell+ " | ");
				   break;
				   }
				System.out.println(" ");
			}
		 }

	}

}
