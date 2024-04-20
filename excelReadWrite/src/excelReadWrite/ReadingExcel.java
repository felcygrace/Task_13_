package excelReadWrite;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcel {

	public static void main(String[] args) {
		ReadingExcel readingexcel = new ReadingExcel();
		readingexcel.readFromExcel();

	}
	public void readFromExcel(){
		//set the stream to connect to excel file
		FileInputStream fis;
		try {
			fis = new FileInputStream("C:\\Users\\FELCY\\eclipse-workspace\\excelReadWrite\\ExcelRead.xlsx");
				//open the workbook
		XSSFWorkbook ExcelWorkbook;
		try {
			ExcelWorkbook = new XSSFWorkbook(fis);
				//open the sheet
		XSSFSheet xlSheet = ExcelWorkbook.getSheetAt(0);
				//get hold of the rows
		 for (Row row : xlSheet) {
             // Iterate through each cell in the row
             for (int c = 0; c < row.getLastCellNum(); c++) {
                 // Print the cell value
                 System.out.print(row.getCell(c) + "\t");
             }
             System.out.println(); // Move to the next line after printing each row
         }
		 ExcelWorkbook.close();
		 fis.close();
		
	}
		 catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		 
		
		
		
	}

}
