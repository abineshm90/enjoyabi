package org.tcs;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class A {
	
	public static void main(String[] args) throws IOException {
		
		
		File loc = new File ("C:\\Users\\srile\\Documents\\New folder\\Newproject\\ExcelNew\\emptynew.xlsx");
		
		FileInputStream stream = new FileInputStream(loc);
						
		Workbook w = new XSSFWorkbook(stream);
		
		Sheet s = w.getSheet("chennai");
		
		Row r = s.getRow(5);
		
		Cell c = r.getCell(6);
		
		String s1 = c.getStringCellValue();
		
		if(s1.equals("srilekha")) {
		
		c.setCellValue("deiva");
		}
		FileOutputStream o= new FileOutputStream(loc);
		
		w.write(o);
		
		System.out.println("wrote sucessfully");
		
		
		
		
		
	}
	
	
}
