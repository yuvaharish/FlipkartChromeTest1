package excelRead;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class ExcelWrite extends ExcelRead{
	
	
	public void excelWrite() throws IOException {
		File file = new File("C:\\Users\\Yuvaraj\\OneDrive\\Desktop\\Attendance Tracker.xlsx");
		FileInputStream fis = new FileInputStream(file);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheet = wb.getSheet("Nov 2022 Batch 2");
		Row createRow = sheet.createRow(1);
		Cell createCell = createRow.createCell(3);
		createCell.setCellValue("Murugan Yuvaraj");
		
		FileOutputStream fo =new FileOutputStream(file);
		wb.write(fo);
		wb.close();
	}
	

	public void updateExcel() throws IOException {
		File file = new File("C:\\Users\\Yuvaraj\\OneDrive\\Desktop\\Attendance Tracker.xlsx");
		FileInputStream fis = new FileInputStream(file);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheet = wb.getSheet("Nov 2022 Batch 2");
		Row row = sheet.getRow(1);
		Cell cell = row.getCell(3);
		String stringCellValue = cell.getStringCellValue();
		System.out.println(stringCellValue);
		if (stringCellValue.equals("Murugan Yuvaraj")) {
			cell.setCellValue("Murugan Senthil");
			FileOutputStream fo = new FileOutputStream(file);
			wb.write(fo);
			wb.close();
		}
		
	}
	@Test
	public void chrome() throws IOException {
		
		System.out.println(ExcelRead.reusableExcel(0, 1));
	}

}
