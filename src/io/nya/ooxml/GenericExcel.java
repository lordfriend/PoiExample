package io.nya.ooxml;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class GenericExcel {
	
	public static final int INVALID_COLUMN_COUNT = -1;
	
	private SXSSFWorkbook mWb;
	
	private Sheet mCurrentSheet;
	private int mCurrentSheetColumnCount = INVALID_COLUMN_COUNT;
	private HashSet<CellRangeAddress> mMergedRanges;
	private int mRowIndex = 0;
	
	private HashMap<String, CellStyle> mSharedCellStyles;
	
	public GenericExcel() {
		init("default");
	}
	
	public GenericExcel(String defaultSheetName) {
		init(defaultSheetName);
	}
	
	private void init(String defaultSheetName) {
		mWb = new SXSSFWorkbook(100);
		addSheet(defaultSheetName);
	}
	
	public void writeLine(ArrayList<CellDefine> rowData) {
		if(mCurrentSheetColumnCount == INVALID_COLUMN_COUNT) {
			mCurrentSheetColumnCount = getColumnCount(rowData);
		}
		int colIndex = 0;
		Row row = null;
		Sheet sheet = mCurrentSheet;
		for(CellDefine cellDefine: rowData) {
			int colToSkip = 0;
			do {
				colToSkip = checkMergedRegion(mRowIndex, colIndex);
				if(colToSkip > 0) {
					colIndex += colToSkip + 1;					
				}
				if(colIndex >= mCurrentSheetColumnCount) {
					mRowIndex++;
					colIndex = 0;
				}
			} while(colToSkip > 0);
			
			if(row == null) {
				row = sheet.createRow(mRowIndex);
			}
			
			createCell(row, mRowIndex, colIndex, cellDefine);
			
			int colSpan = checkAndStoreMergeRegion(cellDefine, mRowIndex, colIndex);
			// if colSpan is greater than 1, we skip the n colSpan to avoid checkMergedRegion
			if(colSpan > 1) {
				colIndex += colSpan;
			}
		}
		mRowIndex++;
		removeOldRanges(mRowIndex);
	}
	
	public void addSheet(String sheetName) {
		mCurrentSheet = mWb.createSheet(sheetName);
		mCurrentSheetColumnCount = INVALID_COLUMN_COUNT;
		mRowIndex = 0;
		mMergedRanges = new HashSet<CellRangeAddress>();
	}
	
	public void addSharedStyle(CellStyleDefine cellStyleDefine) {
		CellStyle style = mWb.createCellStyle();
		
	}
	
	private int getColumnCount(ArrayList<CellDefine> line) {
		int count = 0;
		for(CellDefine cellDefine: line) {
			count += cellDefine.colSpan;
		}
		return count;
	}
	
	private void createCell(Row row, int rowIndex, int colIndex, CellDefine cellDefine) {
		Cell cell = row.createCell(colIndex);
		cell.setCellValue(cellDefine.data);
//		cell.setCellStyle(cellDefine.styleName);
		cell.setCellType(cellDefine.type);
	}
	
	private int checkAndStoreMergeRegion(CellDefine cellDefine, int rowIndex, int colIndex) {
		if(cellDefine.rowSpan > 1 || cellDefine.colSpan > 1) {
			CellRangeAddress rangeAddress = new CellRangeAddress(rowIndex, rowIndex + cellDefine.rowSpan - 1, colIndex, colIndex + cellDefine.colSpan - 1);
			if(cellDefine.rowSpan > 1) {
				mMergedRanges.add(rangeAddress);
			}
			mCurrentSheet.addMergedRegion(rangeAddress);
		}
		return cellDefine.colSpan;
	}
	
	private void removeOldRanges(int rowIndex) {
		Iterator<CellRangeAddress> iterator = mMergedRanges.iterator();
		while(iterator.hasNext()) {
			CellRangeAddress range = iterator.next();
			int lastRow = range.getLastRow();
			if(lastRow < rowIndex) {
				iterator.remove();
			}
		}
	}
	
	/*
	 * We don't use the {@link Sheet#getMergedRegion(index)} method to avoid iterate over all the merged range.
	 * When check whether a cell position is merged other previous cell. only need to store some cross rows range.
	 * The cross columns range is in the incoming CellDefine. 
	 */
	private int checkMergedRegion(int row, int col) {
		boolean[] invalidCellColIndex = new boolean[mCurrentSheetColumnCount - col];
		Arrays.fill(invalidCellColIndex, false);
		for(CellRangeAddress range: mMergedRanges) {
			int firstRow = range.getFirstRow();
			int lastRow = range.getLastRow();
			if(firstRow <= row && lastRow >= row) {
				int firstColumn = range.getFirstColumn();
				int lastColumn = range.getLastColumn();
				if(firstColumn >= col) {
					for(int i = firstColumn - col; i <= lastColumn - col; i++) {
						invalidCellColIndex[i] = true;
					}
				}
			}
		}
		int colToSkipCount = 0;
		for(int i = 0; i < invalidCellColIndex.length && invalidCellColIndex[i]; i++) {
			colToSkipCount++;
		}
		return colToSkipCount;
	}
}
