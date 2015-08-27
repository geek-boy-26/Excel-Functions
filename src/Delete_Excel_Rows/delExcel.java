package Delete_Excel_Rows;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sun.org.apache.bcel.internal.generic.GETSTATIC;

import ReadExcel.ReadExcel;

public class delExcel {

	public static void main(String[] args) 
	{
		// TODO Auto-generated method stub
		try
		{
			FileInputStream file = new FileInputStream(new File("Employee_Details.xlsx"));
			
			//Create Workbook instance 
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			
			//Get Sheet/desired sheet from workbook
			XSSFSheet sheet = workbook.getSheetAt(0);
			
			
			//Iterate Through rows one by one
			Iterator<Row> rowIterator = sheet.iterator();
			while(rowIterator.hasNext())
			{
				Row row = rowIterator.next();
				//For Each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();
		
				while(cellIterator.hasNext())
				{
					Cell cell = cellIterator.next();
					//Check the cell type and format accordingly
					switch(cell.getCellType())
					{
				
					case Cell.CELL_TYPE_NUMERIC:
						System.out.println(cell.getNumericCellValue() + "\t");
						break;
					case Cell.CELL_TYPE_STRING:
						System.out.println(cell.getStringCellValue()+ "\t");
						if(cell.getStringCellValue() == "kabir")
						{
							System.out.println("True"+ ":::");
						}
						break;
					}
				}
				
				System.out.println("");
				
			}
			file.close();
			
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	
		
		
		
	}

}
