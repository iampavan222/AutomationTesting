package Apache_Poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ApplyFormulaOnCode08 
{
    public static FileInputStream fi;
    public static Workbook wb;
    public static Sheet sh;
    public static Row row;
    public static Cell cell;
    public static FileOutputStream fo;
	public static void main(String[] args) throws IOException 
	{
	 String file="D:\\selenium\\Apache Poi\\src\\Data\\CellFormula.xlsx";
	 fi = new FileInputStream(file);
	 wb = new XSSFWorkbook(fi);
	 sh = wb.getSheet("Sheet1");
	 row=sh.getRow(6);
	 cell=row.createCell(2);
	 cell.setCellFormula("SUM(C2:C5)");
	 fo= new FileOutputStream(file);
	 wb.write(fo);
	 wb.close();
  }

}
