package writer;

import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import mapper.ExcelRecord;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import example.SimpleRecord;

public class ExcelWriter {

	public void write(String fileName, XSSFWorkbook workbook) throws Exception {
		FileOutputStream outputStream = new FileOutputStream(fileName);
		workbook.write(outputStream);
		workbook.close();
	}

	public XSSFWorkbook createWorkbook(List<String> headers, List<ExcelRecord> records) {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Sheet1");
		Integer index = 0;
		Row headerRow = sheet.createRow(0);
		for(String header : headers){
			this.createHeader(header, headerRow, index++);
		}
		index = 1;
		for(ExcelRecord excelRecord : records){
			Row row = sheet.createRow(index++);
			excelRecord.getRow(workbook, row);
		}
		return workbook;
	}

	private void createHeader(String header, Row row, Integer index) {
		Cell cell = row.createCell(index);
		cell.setCellValue(header);		
	}

	public static void main(String[] args) throws Exception {
		ExcelWriter excelWriter = new ExcelWriter();
		ExcelRecord simpleRecord = new SimpleRecord("pepito", new Date(), 15.5d);
		List<String> headets = Arrays.asList("nombre", "fecha", "monto");
		XSSFWorkbook workbook = excelWriter.createWorkbook(headets, Arrays.asList(simpleRecord));		
		new ExcelWriter().write("MyFirstExcel.xls", workbook);
	}
}
