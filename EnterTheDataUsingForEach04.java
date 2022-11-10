package Apache_Poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EnterTheDataUsingForEach04 
{
    public static FileInputStream fi;
    public static Workbook wb;
    public static Sheet sh;
    public static Row row;
    public static Cell cell;
    public static FileOutputStream fo;
	public static void main(String[] args) throws IOException 
	{
	String file="D:\\selenium\\Apache Poi\\src\\Data\\sample sheet.xlsx";	
	fi = new FileInputStream(file);
	wb = new XSSFWorkbook(fi);
	sh= wb.createSheet("Emp2");
	/*
	Object empdata [][]= { {"EmpId","Name","Designation","Salary"},
            {"E101","John","QA","20000"},
            {"E102","Romeo","Dev","60000"},
            {"E103","Rolex","BA","50000"},
            };
            */
	ArrayList<Object[]> empdata= new ArrayList<Object[]>();
	empdata.add(new Object[] {"Empid","Name","Desg"});
	empdata.add(new Object[] {"1000","Raju","QA"});
	empdata.add(new Object[] {"1001","Ramu","Dev"});
	empdata.add(new Object[] {"1002","Rohan","QA"});
	int rowcnt=0;
	for(Object emp[]:empdata) 
	{
		row=sh.createRow(rowcnt++);
		 int colcnt=0;
		 for(Object value:emp) 
		 {
			cell=row.createCell(colcnt++); 
			if(value instanceof String)
				cell.setCellValue((String)value);
			if(value instanceof Integer)
				cell.setCellValue((Integer)value);
			if(value instanceof Boolean)
				cell.setCellValue((Boolean)value);
	     }
	}
	
	fo = new FileOutputStream(file);
	wb.write(fo);
	wb.close();
	}

}
