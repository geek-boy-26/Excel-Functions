package WriteExcel;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.hmef.attribute.MAPIAttribute;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import sun.org.mozilla.javascript.internal.ObjArray;



public class WriteExcel 
{

	public static void main(String[] args)
	{
		// TODO Auto-generated method stub
		//Create Blank Workbook
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		//Create Blank Sheet
		XSSFSheet sheet = workbook.createSheet("Employee Data");
		
		//Add data to sheet
		Map<String, Object[]> data = new TreeMap<String, Object[]>();
		data.put("1",new Object[] {"ID","NAME","LASTNAME","E-mail"});
		data.put("2",new Object[] {1,"kabir","khan","kabir.khan@gmail.com"});
		data.put("3",new Object[] {2,"geetesh","chauhan","geet.chauhan@gmail.com"});
		data.put("4",new Object[] {3,"Sachin","Mahajan","sachin.hacker@gmail.com"});
		data.put("5",new Object[] {4,"Anil","Goplani","anil.gops@gmail.com"});
		data.put("6",new Object[] {5,"Bhavesh","parmar","bhav.par@gmail.com"});
		data.put("7",new Object[] {6,"Mahesh","Malsatthar","Mahesh.no17@gmail.com"});
		
		//Iterate over data and write to sheet
		Set<String> keyset = data.keySet();
		int rownum= 0;
		for(String key: keyset)
		{
			Row row = sheet.createRow(rownum++);
			Object[] objArr = data.get(key);
			int cellnum = 0;
			for (Object obj : objArr)
			{
				Cell cell = row.createCell(cellnum++);
				if(obj instanceof String)
					cell.setCellValue((String)obj);
				else if(obj instanceof Integer)
					cell.setCellValue((Integer)obj);
			}
			
			
		}
		try
		{
			//Write in Excel
			FileOutputStream out = new FileOutputStream(new File("Employee_Details.xlsx"));
			workbook.write(out);
			out.close();
			workbook.close();
			System.out.println("Done");
			
			
		}
		catch(Exception e)
		{
			e.printStackTrace();
			e.getMessage();
		}
		
	}

	
	
}
