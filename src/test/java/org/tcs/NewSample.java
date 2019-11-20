package org.tcs;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class NewSample {
	
	public static String getData( int rowNo,int cellNo) throws IOException {
		
		String value= null;
		
		File loc = new File ("C:\\Users\\srile\\Documents\\New folder\\Newproject\\ExcelNew\\excel.xlsx");
		
		FileInputStream stream = new FileInputStream(loc);
		
		Workbook w = new XSSFWorkbook(stream);
		
		Sheet s = w.getSheet("Sheet1");
				
			Row r = s.getRow(rowNo);
			
			Cell c = r.getCell(cellNo);
			
			int type = c.getCellType();
			
			System.out.println(type);
			

			if (type==1) {
				
				value = c.getStringCellValue();
				System.out.println(value);
			}
			
			else if (type==0) {
			
				if (DateUtil.isCellDateFormatted(c)) {
			
				Date gdcv = c.getDateCellValue();
				
				SimpleDateFormat sim = new SimpleDateFormat("dd-MMM-yy");
				
				 value = sim.format(gdcv);
				
						
			} 
			}
				else {
					double n = c.getNumericCellValue();
					long l= (long)n;
					
					 value = String.valueOf(l);
					
				}
			
			return null;
			}
			
	}
			
			
	
			
		
		
		
		
	
