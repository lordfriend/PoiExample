package io.nya.ooxml;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

public class PoiExample {
	
	private static HashMap<String, String> regionMap = new HashMap<String, String>();
	private static HashMap<String, Integer> provinceData = new HashMap<String, Integer>();

	public static void main(String[] args) throws Exception {
		
		regionMap.put("Texas", "US");
		regionMap.put("California", "US");
		regionMap.put("New Jersy", "US");
		regionMap.put("Ohio", "US");
		regionMap.put("Georgia", "US");
		regionMap.put("Beijing", "CN");
		regionMap.put("Shanghai", "CN");
		regionMap.put("Guangdong", "CN");
		regionMap.put("Zhejiang", "CN");
		
		int saleCount = 100;
		int sumCount = 0;
		for(String prov: regionMap.keySet()) {
			provinceData.put(prov, saleCount);
			saleCount+=10;
			sumCount+=saleCount;
		}
		
		ArrayList<CellDefine> firstRow = new ArrayList<CellDefine>();
		firstRow.add(new CellDefine());
		firstRow.add(new CellDefine());
		firstRow.get(0).data = "时间范围";
		firstRow.get(0).rowSpan = 3;
		firstRow.get(0).styleName = "header";
		firstRow.get(1).data = "线下销售";
		firstRow.get(1).colSpan = provinceData.size() + 1;
		firstRow.get(1).styleName = "header";
		
		ArrayList<CellDefine> secondRow = new ArrayList<CellDefine>();
		
		for(String country: regionMap.values()) {
			int index = secondRow.indexOf(country);
			if(index == -1) {
				CellDefine newCell = new CellDefine();
				newCell.data = country;
				newCell.styleName = "header";
				secondRow.add(newCell);
			} else {
				CellDefine cell = secondRow.get(index);
				cell.colSpan++;
			}
		}
		CellDefine sumTitleCell = new CellDefine();
		sumTitleCell.data = "合计";
		sumTitleCell.rowSpan = 2;
		sumTitleCell.styleName = "header";
		secondRow.add(sumTitleCell);
		
		ArrayList<CellDefine> thirdRow = new ArrayList<CellDefine>();
		
		for(CellDefine cell: secondRow) {
			String country = cell.data;
			for(Entry<String, String> entry: regionMap.entrySet()) {
				if(entry.getValue().equals(country)) {
					CellDefine provCell = new CellDefine();
					provCell.data = entry.getKey();
					thirdRow.add(provCell);
				}
			}
		}
		
		ArrayList<CellDefine> forthRow = new ArrayList<CellDefine>();
		CellDefine timeSpanCell = new CellDefine();
		timeSpanCell.data = "20140101-20140909";
		timeSpanCell.styleName = "body";
		
		for(CellDefine cell: thirdRow) {
			String prov = cell.data;
			int saleData = provinceData.get(prov);
			CellDefine saleCell = new CellDefine();
			saleCell.data = String.valueOf(saleData);
			saleCell.type = Cell.CELL_TYPE_NUMERIC;
			saleCell.styleName = "body";
			forthRow.add(saleCell);
		}
		
		CellDefine sumCell = new CellDefine();
		sumCell.data = String.valueOf(sumCount);
		sumCell.type = Cell.CELL_TYPE_NUMERIC;
		sumCell.styleName = "body";
		forthRow.add(sumCell);
		
		CellStyleDefine headerStyle = new CellStyleDefine();
		
		headerStyle.font = new FontDefine();
		headerStyle.font.bold_weight = 400;
		headerStyle.border = new BorderDefine();
		headerStyle.border.style = "THICK";
		headerStyle.fill = new FillDefine();
		headerStyle.fill.background_color = "yellow";
		headerStyle.alignment = "CENTER";
		
		CellStyleDefine bodyStyle = new CellStyleDefine();
		bodyStyle.border = new BorderDefine();
		bodyStyle.border.style = "THIN";
		
		GenericExcel excel = new GenericExcel();
		excel.addSharedStyle("header", headerStyle);
		excel.addSharedStyle("body", bodyStyle);
		
		excel.writeLine(firstRow);
		excel.writeLine(secondRow);
		excel.writeLine(thirdRow);
		excel.writeLine(forthRow);
		
		excel.writeToFile("test.xlsx");
	}

}
