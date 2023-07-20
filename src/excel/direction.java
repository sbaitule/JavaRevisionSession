package excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public interface direction {

//	provide path of Excel Sheet
	String path = "P:\\Snehal\\Velocity\\task.xlsx";
	
	default void input() throws EncryptedDocumentException, IOException {
		
		FileInputStream file = new FileInputStream(path);
		
//		open and read / fetch the excel sheet
	String value = WorkbookFactory.create(file).getSheet("Sheet1").getRow(0).getCell(0).getStringCellValue();
	System.out.println("Value is " + value);
	System.out.println("-------------------------------------");
	
	}
	
	
	default void nextfoermate() throws EncryptedDocumentException, IOException {
		
		 for(int i=0;i<3;i++) {
			 for(int j=0;j<3;j++) {
				 
			 
		String path2 = "P:\\Snehal\\Velocity\\task.xlsx\\";
		FileInputStream file2 = new FileInputStream (path2);
		
		String value2 = WorkbookFactory.create(file2).getSheet("Sheet1").getRow(i).getCell(j).getStringCellValue();
		
		System.out.println(" " +value2);
		 }
		 }
		 System.out.println("-------------------------------------");
	}
	
	default void squareformat() throws EncryptedDocumentException, IOException
	{
		for(int i=0;i<3;i++) {
			for(int j=0;j<3;j++)
			{
				String path3 = "P:\\Snehal\\Velocity\\task.xlsx\\";
				FileInputStream file3 = new FileInputStream(path3);
				String value3 = WorkbookFactory.create(file3).getSheet("Sheet1").getRow(i).getCell(j).getStringCellValue();
				System.out.print("  " +value3);
			}
			System.out.println();
		}
		System.out.println("-------------------------------------");
	}
	
	default void reverceSquare() throws EncryptedDocumentException, IOException
	{
		for (int i = 0;i<3;i++) {
			for(int j=0;j<3;j++) {
				String path4 = "P:\\Snehal\\Velocity\\task.xlsx\\";
				FileInputStream file4 = new FileInputStream(path4);
				String value4 = WorkbookFactory.create(file4).getSheet("Sheet1").getRow(j).getCell(i).getStringCellValue();
				System.out.print("  " + value4);
			}
			System.out.println();
		}
		System.out.println("-------------------------------------");
	}
	
	public default void output() throws EncryptedDocumentException, IOException {
		String way = "P:\\Snehal\\Velocity\\task.xlsx\\";
		FileInputStream f =new FileInputStream(way);
//		double v = WorkbookFactory.create(f).getSheet("Sheet2").getRow(0).getCell(0).getNumericCellValue();
	
//	System.out.println(v);
	
CellType type = WorkbookFactory.create(f).getSheet("Sheet3").getRow(0).getCell(0).getCellType();
	System.out.println(type);
	
	}
}
