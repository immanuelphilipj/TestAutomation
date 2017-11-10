package com.test.excel;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*
 * 
 * Read an Excel file and return in Array of Objects from readExcel method
 * 
 */
public class ExcelReading {

	public static final String path ="C:\\testing\\Book1.xlsx";
	
	public static ArrayList<Object> readExcel(int sheetno, int cellno, int rowno){
		try{
			
			FileInputStream excel = new FileInputStream(new File(path));
			XSSFWorkbook workbook = new XSSFWorkbook(excel);  //Fetch workbook
			XSSFSheet sheet = workbook.getSheetAt(sheetno);  //Fetch Sheet
			Cell cellnumb = sheet.getRow(rowno).getCell(cellno); //Fetch Individual cell
			ArrayList<Object> al = new ArrayList<Object>();
			if(cellnumb.getCellTypeEnum() == CellType.NUMERIC)
			{
				//System.out.println("Numeric valie "+cellnumb.getNumericCellValue());
				al.add(cellnumb.getNumericCellValue());				
			}
			else if(cellnumb.getCellTypeEnum() == CellType.STRING)
			{	
			//System.out.println("Independent cell value "+cellnumb.getStringCellValue());
			al.add(cellnumb.getStringCellValue());
			}
			return al;
/*        To Fetch every values in the excel sheet			
			Iterator<Row> it = sheet.iterator();
			ArrayList<Object> al = new ArrayList<Object>(); 
			
			while(it.hasNext()){
				
				Row rowCell = it.next();
				Iterator<Cell> cell =rowCell.iterator();
				
				while(cell.hasNext()){
					
					Cell currentCell = cell.next();
					if(currentCell.getCellTypeEnum() == CellType.STRING){
					
						al.add(currentCell.getStringCellValue());
					}					
					else if(currentCell.getCellTypeEnum() == CellType.NUMERIC){
						al.add(currentCell.getNumericCellValue());
					}					
				}
				
			}			
		
			return al;  */
		}catch(Exception e){
			System.out.println("Exception is "+e);
			return null;
		}
	}	
}
