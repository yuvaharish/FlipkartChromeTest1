package excelRead;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class ExcelRead {
	

	public static String reusableExcel(int rowNumber,int cellNumber) throws IOException {
		File file = new File("C:\\Users\\Yuvaraj\\OneDrive\\Desktop\\Attendance Tracker.xlsx");
		FileInputStream fis = new FileInputStream(file);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheet = wb.getSheet("Nov 2022 Batch 2");
			Row row = sheet.getRow(rowNumber);
			Cell cell = row.getCell(cellNumber);
			
			int cellType = cell.getCellType();
			String value=null;
				if (cellType==1) {
					value = cell.getStringCellValue();
					
				}else if(cellType==0) {
					if (DateUtil.isCellDateFormatted(cell)) {
						Date dateCellValue = cell.getDateCellValue();
						SimpleDateFormat sd = new SimpleDateFormat("dd/MM/yyyy");
						value = sd.format(dateCellValue);
						
					}else {
						double numericCellValue = cell.getNumericCellValue();
						long l =(long) numericCellValue;
						value = String.valueOf(l);
						
					}
				}
				return value;
	}
	
	@Test
	public void readExcel() throws IOException {
		File file = new File("C:\\Users\\Yuvaraj\\OneDrive\\Desktop\\Attendance Tracker.xlsx");
		FileInputStream fis = new FileInputStream(file);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheet = wb.getSheet("Nov 2022 Batch 2");
		int rowCount = sheet.getPhysicalNumberOfRows();
		
		for (int i = 0; i <rowCount; i++) {
			Row row = sheet.getRow(i);
			int cellCount = row.getPhysicalNumberOfCells();
			for (int j = 0; j < cellCount; j++) {
				Cell cell = row.getCell(j);
				int cellType = cell.getCellType();
				if (cellType==1) {
					String value = cell.getStringCellValue();
					System.out.println(value);
				}else if(cellType==0) {
					if (DateUtil.isCellDateFormatted(cell)) {
						Date dateCellValue = cell.getDateCellValue();
						SimpleDateFormat sd = new SimpleDateFormat("dd/MM/yyyy");
						String value = sd.format(dateCellValue);
						System.out.println(value);
					}else {
						double numericCellValue = cell.getNumericCellValue();
						long l =(long) numericCellValue;
						String value = String.valueOf(l);
						System.out.println(value);
					}
				}
			}
		}
	}
	
public static void main(String[] args) throws IOException {
	String reusableExcel = reusableExcel(0, 2);
	System.out.println(reusableExcel);
}
}
