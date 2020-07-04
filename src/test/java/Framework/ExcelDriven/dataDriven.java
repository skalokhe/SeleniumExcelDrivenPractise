package Framework.ExcelDriven;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDriven {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		//get the file location of excel
		
		FileInputStream fins= new FileInputStream("D:\\Udemy\\MyAutomation\\ExcelDriven\\ExcelData\\Data01.xlsx");
		//access the excel i.e read so create workbook object
		XSSFWorkbook workbook = new XSSFWorkbook(fins);
		
		//to access a particular sheet, get the count of the sheet
		int sheetcount =workbook.getNumberOfSheets();
		// access the sheet
		
		
		for(int i=0;i<sheetcount;i++)
		{
			if(workbook.getSheetName(i).equalsIgnoreCase("DataSheet_01"))
			{
				XSSFSheet sheet=workbook.getSheetAt(i);
				
				//iterate Rows
				Iterator<Row> rows = sheet.iterator();
				Row firstrow= rows.next();
				
				Iterator<Cell> cells=firstrow.cellIterator();
				int k=0,column=0;
				while(cells.hasNext())
				{
					Cell value=cells.next();
					if(value.getStringCellValue().equalsIgnoreCase("testcases"))
					{
						//desired column
						column=k;
					}
					k++;
				}
				System.out.println("Column : "+column);
				
				///iterate thorugh row and get all data
				while(rows.hasNext())
				{
					Row r=rows.next();
					
					if(r.getCell(column).getStringCellValue().equalsIgnoreCase("purchase"))
					{
						Iterator <Cell> c = r.cellIterator();
						while(c.hasNext())
						{
							System.out.println(c.next().getStringCellValue());
						}
					}
				}
			}
		}
		
		

	}

}
