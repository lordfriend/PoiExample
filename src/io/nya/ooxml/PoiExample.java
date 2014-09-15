package io.nya.ooxml;

import io.nya.ooxml.define.BorderDefine;
import io.nya.ooxml.define.CellDefine;
import io.nya.ooxml.define.CellStyleDefine;
import io.nya.ooxml.define.FillDefine;
import io.nya.ooxml.define.FontDefine;
import io.nya.ooxml.pojo.Detail;
import io.nya.ooxml.pojo.Online;
import io.nya.ooxml.pojo.Region;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Cell;

import com.fasterxml.jackson.databind.ObjectMapper;

import redis.clients.jedis.Jedis;
import redis.clients.jedis.JedisPool;
import redis.clients.jedis.JedisPoolConfig;

public class PoiExample {
	
	public static void main(String[] args) throws Exception {
//		 generalDataTest();
//		redisTest();
		System.out.println("Start generate: " + args[1]);
		if(args.length != 3) {
			System.err.println("ilegall arguments");
		} else {
			if(args[0].equals("-region")) {
				generateRegion(args[1], args[2]);
			} else if(args[0].equals("-online") || args[0].equals("-device")) {
				generateOnlineOrDevice(args[0], args[1], args[2]);
			} else if(args[0].equals("-detail")) {
				generateDetail(args[1], args[2]);
			}
		}
	}
	
	private static CellStyleDefine getStyle(String name) {
		if(name.equals("header")) {
			CellStyleDefine headerStyle = new CellStyleDefine();
			
			headerStyle.border = new BorderDefine();
			headerStyle.border.style = "THIN";
			headerStyle.fill = new FillDefine();
			headerStyle.fill.foreground_color = "#FFEDBD";
			headerStyle.fill.fill_pattern = "SOLID_FOREGROUND";
			headerStyle.alignment = "CENTER";
			return headerStyle;
		} else if(name.equals("body")) {
			
			CellStyleDefine bodyStyle = new CellStyleDefine();
			bodyStyle.border = new BorderDefine();
			bodyStyle.border.style = "THIN";
			
			return bodyStyle;
		} else {
			return null;
		}
	}
	
	private static void generateDetail(String exportId, String filename) {
		Jedis jedis = new Jedis("localhost");
		long length = jedis.llen(exportId);
		final long THRESHOLD = 100000;

		String[] titles = new String[]{"序列号", "上传日期", "激活日期", "线上线下", "机型", "所对应的代理商", "省份", "分区"};
		
		GenericExcel excel = new GenericExcel();
		
		excel.addSharedStyle("header", getStyle("header"));
		excel.addSharedStyle("body", getStyle("body"));
		
		ArrayList<CellDefine> firstRow = new ArrayList<CellDefine>();
		
		for(int i = 0; i < titles.length; i++) {
			firstRow.add(new CellDefine(titles[i], "header"));
		}
		
		excel.writeLine(firstRow);
		
		ObjectMapper mapper = new ObjectMapper();
		
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
		
		if(length > THRESHOLD) {
			long start = 0;
			try {
				while(start < length) {
					List<String> result = jedis.lrange(exportId, start, start + THRESHOLD - 1);
					
					for(String rawData: result) {
						Detail detail = mapper.readValue(rawData, Detail.class);
						ArrayList<CellDefine> dataRow = new ArrayList<CellDefine>();
						dataRow.add(new CellDefine(detail.sn, "body"));
						if(detail.upload_time == 0) {
							dataRow.add(new CellDefine("", "body"));
						} else {
							dataRow.add(new CellDefine(sdf.format(new Date(detail.upload_time)), "body"));						
						}
						dataRow.add(new CellDefine(sdf.format(new Date(detail.date)), "body"));
						dataRow.add(new CellDefine(detail.isOnline, "body"));
						dataRow.add(new CellDefine(detail.device, "body"));
						dataRow.add(new CellDefine(detail.orderCompany, "body"));
						dataRow.add(new CellDefine(detail.province, "body"));
						dataRow.add(new CellDefine(detail.region, "body"));
						excel.writeLine(dataRow);
					}
					
					
					start += THRESHOLD;

				}
				
				excel.writeToFile(filename);
				
				System.out.println("done");
			} catch(Exception e) {
				e.printStackTrace();
				System.err.println("failed");
				try {
					excel.writeToFile(filename);
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			} finally {
				jedis.close();
			}
		} else {
			try {
				List<String> result = jedis.lrange(exportId, 0, length);
				
				for(String rawData: result) {
					Detail detail = mapper.readValue(rawData, Detail.class);
					ArrayList<CellDefine> dataRow = new ArrayList<CellDefine>();
					dataRow.add(new CellDefine(detail.sn, "body"));
					if(detail.upload_time == 0) {
						dataRow.add(new CellDefine("", "body"));
					} else {
						dataRow.add(new CellDefine(sdf.format(new Date(detail.upload_time)), "body"));						
					}
					
					dataRow.add(new CellDefine(sdf.format(new Date(detail.date)), "body"));
					dataRow.add(new CellDefine(detail.isOnline, "body"));
					dataRow.add(new CellDefine(detail.device, "body"));
					dataRow.add(new CellDefine(detail.orderCompany, "body"));
					dataRow.add(new CellDefine(detail.province, "body"));
					dataRow.add(new CellDefine(detail.region, "body"));
					excel.writeLine(dataRow);
				}
				

				excel.writeToFile(filename);
				
				System.out.println("done");
			} catch (Exception e) {
				e.printStackTrace();
				System.err.println("failed");
				try {
					excel.writeToFile(filename);
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			} finally {
				jedis.close();
			}
		}
	}

	private static void generateOnlineOrDevice(String type, String exportId, String filename) {
		Jedis jedis = new Jedis("localhost");
		String result = jedis.get(exportId);
		if(result == null) {
			System.err.println("export_id may be expired");
		} else {
			ObjectMapper mapper = new ObjectMapper();
			try {
				Online online = mapper.readValue(result, Online.class);
				
				ArrayList<CellDefine> firstRow = new ArrayList<CellDefine>();
				firstRow.add(new CellDefine());
				firstRow.add(new CellDefine());
				firstRow.get(0).data = "时间范围";
				firstRow.get(0).rowSpan = 2;
				firstRow.get(0).styleName = "header";
				firstRow.get(1).data = type.equals("-online") ? "激活数量（线上机型：通路维度）": "激活数量";
				firstRow.get(1).colSpan = online.content.length + 1;
				firstRow.get(1).styleName = "header";
				
				ArrayList<CellDefine> secondRow = new ArrayList<CellDefine>();
				ArrayList<CellDefine> thirdRow = new ArrayList<CellDefine>();
				
				SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
				CellDefine timeSpanCell = new CellDefine();
				timeSpanCell.data = sdf.format(new Date(online.earliest)) + "~" + sdf.format(new Date(online.latest));
				timeSpanCell.styleName = "body";
				thirdRow.add(timeSpanCell);
				
				int sumCount = 0;

				for(int j = 0; j < online.content.length; j++) {
					CellDefine onlineCell = new CellDefine();
					onlineCell.data = online.content[j]._id;
					onlineCell.styleName = "header";
					CellDefine dataCell = new CellDefine();
					dataCell.data = String.valueOf(online.content[j].total);
					dataCell.type = Cell.CELL_TYPE_NUMERIC;
					dataCell.styleName = "body";
					sumCount += online.content[j].total;
					secondRow.add(onlineCell);
					thirdRow.add(dataCell);
				}
				
				CellDefine sumTitleCell = new CellDefine();
				sumTitleCell.data = type.equals("-online") ? "线上合计" : "当前激活合计";
				sumTitleCell.styleName = "header";
				secondRow.add(sumTitleCell);
				
				CellDefine sumCell = new CellDefine();
				sumCell.data = String.valueOf(sumCount);
				sumCell.type = Cell.CELL_TYPE_NUMERIC;
				sumCell.styleName = "body";
				thirdRow.add(sumCell);
				
				GenericExcel excel = new GenericExcel();
				excel.addSharedStyle("header", getStyle("header"));
				excel.addSharedStyle("body", getStyle("body"));
				
				excel.writeLine(firstRow);
				excel.writeLine(secondRow);
				excel.writeLine(thirdRow);

				excel.autoSizeColumn(0);
				
				excel.writeToFile(filename);
				
				System.out.println("done");
				
			} catch (Exception e) {
				e.printStackTrace();
				System.err.println("fail to parse json");
			} finally {
				jedis.close();
			}
		}
	}

	private static void generateRegion(String exportId, String filename) {
		Jedis jedis = new Jedis("localhost");
		String result = jedis.get(exportId);
		if(result == null) {
			System.err.println("export_id may be expired");
		} else {
			ObjectMapper mapper = new ObjectMapper();
			try {
				Region region = mapper.readValue(result, Region.class);
				
				ArrayList<CellDefine> firstRow = new ArrayList<CellDefine>();
				firstRow.add(new CellDefine());
				firstRow.add(new CellDefine());
				firstRow.get(0).data = "时间范围";
				firstRow.get(0).rowSpan = 3;
				firstRow.get(0).styleName = "header";
				firstRow.get(1).data = "激活数量（线下机型：省/分区维度）";
				firstRow.get(1).colSpan = region.content.length + 1;
				firstRow.get(1).styleName = "header";
				
				ArrayList<String> regionList = new ArrayList<String>();
				for(int i = 0; i < region.content.length; i++) {
					if(!regionList.contains(region.content[i].region)) {
						regionList.add(region.content[i].region);
					}
				}
				
				ArrayList<CellDefine> secondRow = new ArrayList<CellDefine>();
				ArrayList<CellDefine> thirdRow = new ArrayList<CellDefine>();
				ArrayList<CellDefine> forthRow = new ArrayList<CellDefine>();
				
				SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
				CellDefine timeSpanCell = new CellDefine();
				timeSpanCell.data = sdf.format(new Date(region.earliest)) + "~" + sdf.format(new Date(region.latest));
				timeSpanCell.styleName = "body";
				forthRow.add(timeSpanCell);
				
				int sumCount = 0;
				for(int i = 0; i < regionList.size(); i++) {
					CellDefine regionCell = new CellDefine();
					regionCell.data = regionList.get(i);
					regionCell.styleName = "header";
					regionCell.colSpan = 0;
					for(int j = 0; j < region.content.length; j++) {
						if(regionList.get(i).equals(region.content[j].region)) {
							regionCell.colSpan++;
							CellDefine provCell = new CellDefine();
							provCell.data = region.content[j]._id;
							provCell.styleName = "header";
							CellDefine dataCell = new CellDefine();
							dataCell.data = String.valueOf(region.content[j].total);
							dataCell.type = Cell.CELL_TYPE_NUMERIC;
							dataCell.styleName = "body";
							sumCount += region.content[j].total;
							thirdRow.add(provCell);
							forthRow.add(dataCell);
						}
					}
					secondRow.add(regionCell);
				}
				
				CellDefine sumTitleCell = new CellDefine();
				sumTitleCell.data = "合计";
				sumTitleCell.rowSpan = 2;
				sumTitleCell.styleName = "header";
				secondRow.add(sumTitleCell);
				
				CellDefine sumCell = new CellDefine();
				sumCell.data = String.valueOf(sumCount);
				sumCell.type = Cell.CELL_TYPE_NUMERIC;
				sumCell.styleName = "body";
				forthRow.add(sumCell);
				
				GenericExcel excel = new GenericExcel();
				excel.addSharedStyle("header", getStyle("header"));
				excel.addSharedStyle("body", getStyle("body"));
				
				excel.writeLine(firstRow);
				excel.writeLine(secondRow);
				excel.writeLine(thirdRow);
				excel.writeLine(forthRow);

				excel.autoSizeColumn(0);
				
				excel.writeToFile(filename);
				
				System.out.println("done");
				
			} catch (Exception e) {
				e.printStackTrace();
				System.err.println("fail to parse json");
			} finally {
				jedis.close();
			}
		}
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
		ObjectMapper mapper = new ObjectMapper();
		
		GenericExcel excel = new GenericExcel();

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
		
		excel.addSharedStyle("header", headerStyle);
		excel.addSharedStyle("body", bodyStyle);
		
		
		String[] titles = new String[]{"序列号", "激活日期", "线上线下", "设备", "订货单位", "省份", "区域"};
		ArrayList<CellDefine> titleRow = new ArrayList<CellDefine>();
		for(int i = 0; i < titles.length; i++) {
			CellDefine cell = new CellDefine();
			cell.data = titles[i];
			cell.styleName = "header";
			titleRow.add(cell);
		}
		
		excel.writeLine(titleRow);
		
		for(String data: rawList) {
			Detail obj = mapper.readValue(data, Detail.class);
			ArrayList<CellDefine> line = new ArrayList<CellDefine>();
			for(int i = 0; i < titles.length; i++) {
				CellDefine cell = new CellDefine();
				cell.styleName = "body";
				line.add(cell);
			}
			line.get(0).data = obj.sn;
			line.get(1).data = String.valueOf(obj.date);
			line.get(2).data = obj.isOnline;
			line.get(3).data = obj.device;
			line.get(4).data = obj.orderCompany;
			line.get(5).data = obj.province;
			line.get(6).data = obj.region;
			
			excel.writeLine(line);
		}
		
		excel.writeToFile("redis.xlsx");
		
	}

}
