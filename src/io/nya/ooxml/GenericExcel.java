package io.nya.ooxml;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class GenericExcel {
	
	public static final int INVALID_COLUMN_COUNT = -1;
	
	private SXSSFWorkbook mWb;
	
	private Sheet mCurrentSheet;
	private int mCurrentSheetColumnCount = INVALID_COLUMN_COUNT;
	private HashMap<String, HashSet<CellRangeAddress>> mMergedRanges;
	
	public GenericExcel() {
		init("default");
	}
	
	public GenericExcel(String defaultSheetName) {
		init(defaultSheetName);
	}
	
	private void init(String defaultSheetName) {
		mWb = new SXSSFWorkbook(100);
		mCurrentSheet = mWb.createSheet(defaultSheetName);
	}
	
	public void writeLine(ArrayList<CellDefine> row) {
		if(mCurrentSheetColumnCount == INVALID_COLUMN_COUNT) {
			mCurrentSheetColumnCount = getColumnCount(row);
		}
		writeLine(mCurrentSheet, row);
	}
	
	public void writeLine(String sheetName, ArrayList<CellDefine> row) {
		Sheet sheet = mWb.getSheet(sheetName);
		if(sheet == null) {
			sheet = mWb.createSheet(sheetName);
		}
		writeLine(sheet, row);
	}
	
	public boolean changeSheet(String sheetName) {
		mCurrentSheet = mWb.getSheet(sheetName);
		if(mCurrentSheet != null) {
			mCurrentSheetColumnCount = INVALID_COLUMN_COUNT;
			return false;
		} else {
			return false;
		}
	}
	
	public void addSheet(String sheetName) {
		mCurrentSheet = mWb.createSheet(sheetName);
		mCurrentSheetColumnCount = INVALID_COLUMN_COUNT;
	}
	
	private int getColumnCount(ArrayList<CellDefine> line) {
		int count = 0;
		for(CellDefine cellDefine: line) {
			count += cellDefine.colSpan;
		}
		return count;
	}
	
	/*
	 * We don't use the getMergedRegion(int index) method to avoid iterate over all the merged range.
	 * When check whether a cell position is merged other previous cell. only need to store some cross rows range.
	 * The cross columns range is in the incoming CellDefine. 
	 */
	private boolean checkMergedRegion(Sheet sheet, int row, int col) {
		return false;
	}
	
	private void writeLine(Sheet sheet, ArrayList<CellDefine> row) {
		int colnum = 0;
		for(CellDefine cellDefine: row) {
			
		}
	}
}
