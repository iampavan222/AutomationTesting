package Apache_Poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UsingFormulaAndStoreResult09 
{

	public static Workbook wb;
	public static Sheet sh; 
	public static Row row;
	public static Cell cell;
	public static FileOutputStream fo;
	public static void main(String[] args) throws IOException 
	{
		wb = new XSSFWorkbook();
		sh = wb.createSheet("Numbers");
		row=sh.createRow(0);
		row.createCell(0).setCellValue(10);
		row.createCell(1).setCellValue(20);
		row.createCell(2).setCellValue(30);
		row.createCell(3).setCellValue(40);
		
		row.createCell(4).setCellFormula("A1*B1*C1*D1");
		
		fo= new FileOutputStream("D:\\selenium\\Apache Poi\\src\\Data\\Formula.xlsx");
		wb.write(fo);
		wb.close();
		System.out.println("Successfully created formula");
	}

}
