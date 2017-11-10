package com.test.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TestWrite {
	
	private static final String FILE_NAME = "C:\\testing\\Book1.xlsx";
	
	void getExcelData(){
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Data types in Java");
		Object[][] datatypes ={
				{"DataTypes","Types","Size"},
				{"int", "Primitives", 2},
				{"Char","Primitives", 1}
		};
		
		int rownum = 0;
		for(Object[] datatype : datatypes){
			Row row = sheet.createRow(rownum++);
			int colnum = 0;
			for(Object field : datatypes){
				Cell cell = row.createCell(colnum++);
				if(field instanceof String)
					cell.setCellValue( (String) field);
				else if(field instanceof Integer)
					cell.setCellValue((Integer)field);
			}
			
		}
	}
	
	public static void main(String[] args){
		
		try{
			
		}catch(Exception e){
			System.out.println(e);
		}
		

		

	}

}
