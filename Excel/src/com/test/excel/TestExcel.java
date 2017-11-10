package com.test.excel;

import java.util.ArrayList;
import java.util.Iterator;

public class TestExcel {

	public static void main(String[] args){
		
	ArrayList<Object> ab =	ExcelReading.readExcel(0,1,0);
	Iterator<Object> iterator = ab.iterator();
	while(iterator.hasNext()){
		
		System.out.println("Inside "+iterator.next());
	}
		
	}
	
}

