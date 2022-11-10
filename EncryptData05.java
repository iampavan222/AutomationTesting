package Apache_Poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class EncryptData05 
{
	public static FileInputStream fi;
	public static Workbook wb;
	public static Sheet sh;
	public static Row row;
	public static Cell col;
	public static void main(String[] args) throws IOException 
	{
	String path="D:\\selenium\\Apache Poi\\src\\Data\\Encrypt.xlsx";
	String pwd="Pavan1234";
	 fi= new FileInputStream(path);
	 wb=WorkbookFactory.create(fi, pwd);
	 sh=wb.getSheetAt(0);
	 int rw=sh.getLastRowNum();
	 System.out.println(rw);
	 int cl=sh.getRow(1).getLastCellNum();
	 System.out.println(cl);
	 
	 try 
	 {
	 for(int r=0;r<=rw;r++) 
	 {
		 row=sh.getRow(r);
		 for(int c=0;c<cl;c++) 
		 {
			col=row.getCell(c); 
			switch(col.getCellType()) 
			{
			case STRING:
				System.out.print(col.getStringCellValue());
				break;
			case NUMERIC:
				System.out.print(col.getNumericCellValue());
				break;
			case BOOLEAN:
				System.out.print(col.getBooleanCellValue());
				break;
			}
			System.out.print("|");
		 }
		 System.out.println();
        }
	 }
	 catch(Exception E) {
		 System.out.println("test pass");
	 }
	 wb.close();
	 fi.close();
	}

}
