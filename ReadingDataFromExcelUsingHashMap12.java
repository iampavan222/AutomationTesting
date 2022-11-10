package Apache_Poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingDataFromExcelUsingHashMap12 
{
	public static FileInputStream fi;
    public static Workbook wb;
    public static Sheet sh;
    public static Row row;
	public static void main(String[] args) throws IOException 
	{
		String file="D:\\selenium\\Apache Poi\\src\\Data\\SampleData.xlsx";
		fi= new FileInputStream(file);
		wb=new XSSFWorkbook(fi);
		sh=wb.getSheetAt(1);
		int rs=sh.getLastRowNum();
		
		HashMap<String, String> data= new HashMap<String,String>();
		for(int r=0;r<=rs;r++) {
			String key=sh.getRow(r).getCell(0).getStringCellValue();
			String value=sh.getRow(r).getCell(1).getStringCellValue();
			data.put(key, value);
		}
		for(Map.Entry entry:data.entrySet())
		{
			System.out.println(entry.getKey()+ " "+entry.getValue());
			
		}
		System.out.println("task is completed............");
		wb.close();
	}

}
