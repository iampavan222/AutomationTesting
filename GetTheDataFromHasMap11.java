package Apache_Poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GetTheDataFromHasMap11 
{

	public static FileInputStream fi;
	public static Workbook wb;
	public static Sheet sh;
	public static Row row;
	public static  FileOutputStream fo;
	public static void main(String[] args) throws IOException 
	{
		String file="D:\\selenium\\Apache Poi\\src\\Data\\SampleData.xlsx";
		fi=new FileInputStream(file);
	    wb= new XSSFWorkbook(fi);
	    sh=wb.createSheet("HashMap1");
	    
	   Map<String,String> data= new HashMap<String,String>();
	   data.put("101", "Shyam");
	   data.put("102", "Ramu");
	   data.put("103", "Rohit");
	   
	   int rowcount=0;
	   for(Map.Entry entry:data.entrySet()) {
		   row=sh.createRow(rowcount++);
		   row.createCell(0).setCellValue((String)entry.getKey());
		   row.createCell(1).setCellValue((String)entry.getValue());
	   }
	   fo=new FileOutputStream(file);
	   wb.write(fo);
	   wb.close();
}

}
