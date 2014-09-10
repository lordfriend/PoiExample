package io.nya.ooxml;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import redis.clients.jedis.Jedis;
import redis.clients.jedis.JedisPool;
import redis.clients.jedis.JedisPoolConfig;

public class PoiExample {
	
	public static void main(String[] args) throws Exception {
		 generalDataTest();
	}
	
	public static void generalDataTest() throws IOException {
		HashMap<String, String> regionMap = new HashMap<String, String>();
		HashMap<String, Integer> provinceData = new HashMap<String, Integer>();
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
		
		CellDefine cnCell = new CellDefine();
		cnCell.data = "CN";
		cnCell.colSpan = 4;
		cnCell.styleName = "header";
		secondRow.add(cnCell);
		CellDefine usCell = new CellDefine();
		usCell.data = "US";
		usCell.colSpan = 5;
		usCell.styleName = "header";
		secondRow.add(usCell);

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
					provCell.styleName = "header";
					provCell.data = entry.getKey();
					thirdRow.add(provCell);
				}
			}
		}
		
		ArrayList<CellDefine> forthRow = new ArrayList<CellDefine>();
		CellDefine timeSpanCell = new CellDefine();
		timeSpanCell.data = "20140101-20140909";
		timeSpanCell.styleName = "body";
		forthRow.add(timeSpanCell);
		
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
		headerStyle.font.bold_weight = 100;
		headerStyle.border = new BorderDefine();
		headerStyle.border.style = "THICK";
		headerStyle.fill = new FillDefine();
		headerStyle.fill.foreground_color = "yellow";
		headerStyle.fill.fill_pattern = "SOLID_FOREGROUND";
		headerStyle.alignment = "CENTER";
		
		CellStyleDefine bodyStyle = new CellStyleDefine();
		bodyStyle.border = new BorderDefine();
		bodyStyle.border.style = "THIN";
		
		GenericExcel excel = new GenericExcel();
		excel.addSharedStyle("header", headerStyle);
		excel.addSharedStyle("body", bodyStyle);
		
		for(int i = 0; i < firstRow.size(); i++) {
			System.out.println(i+": " + firstRow.get(i).data);			
		}
		for(int i = 0; i < secondRow.size(); i++) {
			System.out.println(i+": " + secondRow.get(i).data);			
		}
		for(int i = 0; i < thirdRow.size(); i++) {
			System.out.println(i+": " + thirdRow.get(i).data);			
		}
		for(int i = 0; i < forthRow.size(); i++) {
			System.out.println(i+": " + forthRow.get(i).data);			
		}
		
		
		excel.writeLine(firstRow);
		excel.writeLine(secondRow);
		excel.writeLine(thirdRow);
		excel.writeLine(forthRow);
		
		for(int i = 0; i < 1000; i++) {
			ArrayList<CellDefine> newRow = new ArrayList<CellDefine>();
			CellDefine timeSpan = new CellDefine();
			timeSpan.data = "20140101-20140909";
			timeSpan.styleName = "body";
			newRow.add(timeSpan);
			int total = 0;
			for(int j = 0; j < provinceData.size(); j++) {
				CellDefine dataCell = new CellDefine();
				dataCell.data = String.valueOf(i + j);
				dataCell.type = Cell.CELL_TYPE_NUMERIC;
				dataCell.styleName = "body";
				newRow.add(dataCell);
				total += i + j;				
			}
			CellDefine totalCell = new CellDefine();
			totalCell.data = String.valueOf(total);
			totalCell.type = Cell.CELL_TYPE_NUMERIC;
			totalCell.styleName = "body";
			newRow.add(totalCell);
			excel.writeLine(newRow);
		}
		
//		excel.setColumnWidth(0, 20);
		excel.autoSizeColumn(0);
		
		excel.writeToFile("test.xlsx");
	}
	
	public static void redisTest() throws IOException {
		JedisPool pool = new JedisPool(new JedisPoolConfig(), "localhost");
		Jedis jedis = null;
		try {
			jedis = pool.getResource();
		} catch(Exception e) {
			e.printStackTrace();
			return;
		}
		String exportId = new BufferedReader(new InputStreamReader(System.in)).readLine();
		long length = jedis.llen(exportId);
		List<String> rawList = jedis.lrange(exportId, 0, length);
		for(String data: rawList) {
			
		}
	}

}
