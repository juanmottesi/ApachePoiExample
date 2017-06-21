package example;

import java.util.Date;

import mapper.ExcelRecord;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class SimpleRecord extends ExcelRecord {

	private String name;
	private Date date;
	private Double number;

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public Date getDate() {
		return date;
	}

	public void setDate(Date date) {
		this.date = date;
	}

	public Double getNumber() {
		return number;
	}

	public void setNumber(Double number) {
		this.number = number;
	}

	public SimpleRecord(String name, Date date, Double number) {
		super();
		this.name = name;
		this.date = date;
		this.number = number;
	}

	public void getRow(Row row) {
		int colNum = 0;
		Cell cell = row.createCell(colNum++);
		cell.setCellValue(this.name);
		cell = row.createCell(colNum++);
		cell.setCellValue(this.date);
		cell.setCellStyle(this.getDateStyle("yyyy-mm-dd"));
		cell = row.createCell(colNum++);
		this.setCellValue(cell, this.number);
	}

}
