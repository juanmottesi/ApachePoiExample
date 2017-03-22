package mapper;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public abstract class ExcelRecord {

	private XSSFWorkbook workbook;

	public abstract void getRow(Row row);

	public void getRow(XSSFWorkbook workbook, Row row) {
		this.workbook = workbook;
		this.getRow(row);
	}

	public CellStyle getDateStyle(String pattern) {
		CellStyle cellStyle = workbook.createCellStyle();
		CreationHelper createHelper = workbook.getCreationHelper();
		cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(
				pattern));
		return cellStyle;
	}
	
	public void setCellValue(Cell cell, Double number){
		if (number == null){
			cell.setCellValue(0);
		} else{
			cell.setCellValue(number);
		}
	}

}
