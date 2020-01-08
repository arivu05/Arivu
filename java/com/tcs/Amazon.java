package com.tcs;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Amazon {
	
	public static void main(String[] args) throws IOException  {
		File loc = new File("C:\\Users\\HP\\eclipse-workspace\\Amazon\\ExcelRead\\sheet2.xlsx");
		Workbook w=new XSSFWorkbook();
		Sheet s = w.createSheet("greens");
		Row r = s.createRow(0);
		Cell c = r.createCell(1);
		c.setCellValue("Arivoli");
		FileOutputStream stream = new FileOutputStream(loc);
		w.write(stream);
		System.out.println("write successfully");
		
		
		
		/**for (int i=0;i<s.getPhysicalNumberOfRows();i++) {
			Row r = s.getRow(i);
			 
			for (int j=0;j<r.getPhysicalNumberOfCells();j++) {
				Cell c = r.getCell(j);
				System.out.println(c);
			}
			
		}**/
	
		
		
	}

}