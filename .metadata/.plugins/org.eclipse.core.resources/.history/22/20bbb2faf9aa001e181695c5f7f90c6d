package com.excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EX_Imp {
	public static void main(String[] args) throws IOException {
		FileInputStream file=new FileInputStream("C:\\Users\\DeepPise\\Desktop\\API\\Dp\\src\\test\\resources\\Book1.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(file);
		int ns = wb.getNumberOfSheets();
		for(int i=0;i<ns;i++)
		{
			if(wb.getSheetName(i).equalsIgnoreCase("Sheet1"))
			{
				XSSFSheet sheet = wb.getSheetAt(i);
				Iterator<Row> rows = sheet.iterator();
				Row firstrow = rows.next();
				Iterator<Cell> ce = firstrow.cellIterator();
				int coloumn=0;
				int k=0;
				while(ce.hasNext())
				{
					Cell value = ce.next();
					if(value.getStringCellValue().equalsIgnoreCase("testcase"))
					{
						coloumn=k;
					}
					k++;
					
				}
				System.out.println(coloumn);
				while(rows.hasNext())
				{
					Row value = rows.next();
					int k1=0;
					if(value.getCell(coloumn).getStringCellValue().equalsIgnoreCase("purchase"))
					{
						Iterator<Cell> cv = value.cellIterator();
						while(cv.hasNext())
						{
							System.out.println(cv.next().getStringCellValue());
							
						}
					}
				}
			}
		}
	}
}
