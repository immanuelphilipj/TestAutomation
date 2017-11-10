package com.test.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingFile {

	boolean t;
	Boolean ts;
	byte b;
	Byte d;
	private static final String FILE_NAME = "C:\\testing\\Book1.xlsx";
	
	public static void main(String args[]) throws Exception{
		ReadingFile rf = new ReadingFile();
	//	rf.myExcel();

		
	}
	
	public void myExcel() throws Exception{
		
		FileInputStream excelFile;
		Workbook workbook = null;
		Sheet datatypeSheet;
		
		try{
		
			File file = new File(FILE_NAME);
            excelFile = new FileInputStream(file);
            workbook = new XSSFWorkbook(excelFile);
            datatypeSheet = workbook.getSheetAt(0);
            //System.out.println("excelFile "+excelFile +"workbook "+workbook+"datatypeSheet "+datatypeSheet);
            
            Iterator<Row> iterator = datatypeSheet.iterator();

            ArrayList<Object> al = new ArrayList<Object>();
            
            while (iterator.hasNext()) {

                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();

                while (cellIterator.hasNext()) {

                    Cell currentCell = cellIterator.next();
                    
                    if (currentCell.getCellTypeEnum() == CellType.STRING) {
                        al.add(currentCell.getStringCellValue());                   
                    }
                    else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                    	al.add(currentCell.getNumericCellValue());                    	                   
                    } 
                }
            }
            for(Object x:al)
                 System.out.println(x);
            
        } catch (Exception e) {
            e.printStackTrace();
        } finally{
        	workbook.close();
        }

}
	}
