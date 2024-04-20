package excelReadWrite;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteToExcel {

    public static void main(String[] args) {
        WriteToExcel writer = new WriteToExcel();
        writer.writeToExcl();
    }

    public void writeToExcl() {
        FileOutputStream fos = null;
        XSSFWorkbook xlWorkbook = null;
        try {
            FileInputStream fis = new FileInputStream("C:\\Users\\FELCY\\eclipse-workspace\\excelReadWrite\\ExcelWrite.xlsx");
            xlWorkbook = new XSSFWorkbook(fis);
            XSSFSheet xlSheet = xlWorkbook.getSheetAt(0);

            XSSFRow xlRow = xlSheet.createRow(0);
            XSSFCell xlCell = xlRow.createCell(0);
            xlCell.setCellValue("Name");
            xlCell = xlRow.createCell(1);
            xlCell.setCellValue("Age");
            xlCell = xlRow.createCell(2);
            xlCell.setCellValue("Email");

            xlRow = xlSheet.createRow(1);
            xlCell = xlRow.createCell(0);
            xlCell.setCellValue("John Doe");
            xlCell = xlRow.createCell(1);
            xlCell.setCellValue("30");
            xlCell = xlRow.createCell(2);
            xlCell.setCellValue("john@Test.com");

            xlRow = xlSheet.createRow(2);
            xlCell = xlRow.createCell(0);
            xlCell.setCellValue("Jane Doe");
            xlCell = xlRow.createCell(1);
            xlCell.setCellValue("28");
            xlCell = xlRow.createCell(2);
            xlCell.setCellValue("john@Test.com");
            
            xlRow = xlSheet.createRow(3);
            xlCell = xlRow.createCell(0);
            xlCell.setCellValue("Bob Smith");
            xlCell = xlRow.createCell(1);
            xlCell.setCellValue("35");
            xlCell = xlRow.createCell(2);
            xlCell.setCellValue("jacky@example.com");
            
            xlRow = xlSheet.createRow(4);
            xlCell = xlRow.createCell(0);
            xlCell.setCellValue("Swapnil");
            xlCell = xlRow.createCell(1);
            xlCell.setCellValue("37");
            xlCell = xlRow.createCell(2);
            xlCell.setCellValue("swapnil@example.com");

            fos = new FileOutputStream("C:\\Users\\FELCY\\eclipse-workspace\\excelReadWrite\\ExcelWrite.xlsx");
            xlWorkbook.write(fos);
            System.out.println("Writing to Excel completed");
            xlWorkbook.close();
            fis.close();
            fos.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } 
        }
    }

