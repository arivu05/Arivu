package com.tcs;

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

public class Flipkart {

	public static void main(String[] args) throws IOException {
		File loc = new File("C:\\\\Users\\\\HP\\\\eclipse-workspace\\\\Amazon\\\\ExcelRead\\\\sheet1.xlsx");
		FileInputStream stream = new FileInputStream(loc);
		Workbook w = new XSSFWorkbook(stream);	
		Sheet s = w.getSheet("sheet1");
		Row r = s.getRow(0);
		Cell c = r.getCell(0);
		int type = c.getCellType();
		//System.out.println(type);
		
		
		
		for (int i=0;i<s.getPhysicalNumberOfRows();i++) {
			s.getRow(i);
			for (int j=0;j<r.getPhysicalNumberOfCells();j++) {
				Cell cell = r.getCell(j);
				//System.out.println(cell);
				
				if (type==1) {
					c.getStringCellValue();
				}
				else if (type==0) {
					
					if (DateUtil.isCellDateFormatted(cell)) {
						Date dateCellValue = cell.getDateCellValue();
						
					//coverting date into cell
						SimpleDateFormat sim = new SimpleDateFormat("dd-mmm-yy");
						String f = sim.format(dateCellValue);
						System.out.println(f);
					}
					else {//check numeric value
						double numericCellValue = cell.getNumericCellValue();
						
						//type cast converting double into long
						long l = (long)numericCellValue;
						
						//converting long to cell
							String valueOf = String.valueOf(j);
							System.out.println(valueOf);
						
					}
				}
			}
		}
		
	}
	
}
