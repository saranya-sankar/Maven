package com.obsqura.mavenProject;
import java.io.*;
import org.apache.poi.xssf.usermodel.*;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
public class ExcelRead {
	public static void main(String[] args) {
		try {
			File file=new File("C:\\Users\\Lenovo\\Desktop\\SampleExcel.xlsx");
			FileInputStream fis=new FileInputStream(file);
			XSSFWorkbook wb=new XSSFWorkbook(fis);
			XSSFSheet sheet=wb.getSheetAt(0);
			Iterator<Row> itr=sheet.iterator();
			while(itr.hasNext()) {
			Row row=itr.next();
			Iterator<Cell> cellIterator=row.cellIterator();
			Cell cell=cellIterator.next();
			switch(cell.getCellTypeEnum()){
			case STRING:
			System.out.println(cell.getStringCellValue()+"\t\t\t");
			break;
			case NUMERIC:
			System.out.println(cell.getNumericCellValue()+"\t\t\t");
			break;
			default:
			}
			
			
		
		System.out.println("");
		//System.out.println("");
			}
		}
	catch(Exception e) {
		e.printStackTrace();
		}
	}
}
