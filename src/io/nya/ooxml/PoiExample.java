package io.nya.ooxml;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class PoiExample {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		SXSSFWorkbook wb = new SXSSFWorkbook(100);
		Sheet sh = wb.createSheet();
		for(int rownum = 0; rownum < 1000; rownum++) {
			Row row = sh.createRow(rownum);
			for(int cellnum = 0; cellnum < 10; cellnum++) {
				Cell cell = row.createCell(cellnum);
				String address = new CellReference(cell).formatAsString();
				cell.setCellValue(address);
			}
		}
		
		FileOutputStream out = new FileOutputStream("sxssf.xlsx");
		wb.write(out);
		out.close();
		wb.dispose();
	}

}
