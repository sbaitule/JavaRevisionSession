package excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class NewExcel {
	public void main (String[]arg){
		
	    int[] school = new int[10];  // for column 0
		for(int i=0;i<=school.length;i++) {
			school[i]= i+1;
		}
		
		String[] name = new String[10]; // for column 1
		name[0]="student A";
		name[1]="student B";
		name[2]="student C";
		name[3]="student D";
		name[4]="student E";
		name[5]="student F";
		name[6]="student G";
		name[7]="student H";
		name[8]="student I";
		name[9]="student J";
		
		String[] result = new String[10];  //for column 2
		result[0] = "pass";
		result[1] = "pass";
		result[2] = "fail";
		result[3] = "pass";
		result[4] = "fail";
		result[5] = "fail";
		result[6] = "pass";
		result[7] = "pass";
		result[8] = "pass";
		result[9] = "pass";
		
		//creat workbook
		XSSFWorkbook wb = new XSSFWorkbook();
		
		//creat spredsheet
		XSSFSheet sheet = new wb.createSheet("velocity");
		
		//creat row
		XSSFSheet row;
		row = new sheet.createRow(0);
		
		Cell cell0 = row.createRow(0);
		Cell cell1 = row.createRow(1);
		Cell cell2 = row.createRow(2);
		
		
		//logic
		
		for(int i =0;i<school.length;i++) {
			row = sheet.createrow(i+1);
			
			for(int j=0;j<=result.length;j++) {
				
				Cell cell = row.createCell(j);
				
				if (cell.getColumnIndex()==0)
				{
					cell.setCellValue(school[i]);
				}
				else if (cell.getColumnIndex()==1)
				{
				    cell.setCellValue(name[i]);
				}
				else if (cell.getColumnIndex()==2)
				{
					cell.setCellValue(result[i]);
				}
			}
			
		}
		String path = "P:\\Snehal\\Velocity\\task1.xlsx";
		try {
			FileOutputStream out = new FileOutputStream(path);
			wb.write(out);
			System.out.println("file generated.............");
			out.close();
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		finally {
			
		}
		
		}
	}


