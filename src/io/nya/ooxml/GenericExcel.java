package io.nya.ooxml;

import java.awt.Color;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FontFamily;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;

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
		short height = -1;
		Sheet sheet = mCurrentSheet;
		// try to get row at index of mRowIndex where we may already create an row for some merged regions.
		Row row = sheet.getRow(mRowIndex);
		if(row != null) {
			height = row.getHeight();
		}
		for(CellDefine cellDefine: rowData) {
			int colToSkip = 0;
			do {
				colToSkip = checkMergedRegion(mRowIndex, colIndex);
//				System.out.println("colToSkip: " + colToSkip);
				if(colToSkip > 0) {
					colIndex += colToSkip;					
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
			
			if(cellDefine.height > -1 && cellDefine.height > height) {
				height = cellDefine.height;
			}
			int colSpan = checkAndStoreMergeRegion(cellDefine, mRowIndex, colIndex);
			// if colSpan is greater than 1, we skip the n colSpan to avoid checkMergedRegion
			if(colSpan > 1) {
				colIndex += colSpan;
			} else {
				colIndex++;
			}
		}
		
		if(height > -1) {
			row.setHeight(height);
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
	
	public void addSharedStyle(String styleName, CellStyleDefine cellStyleDefine) {
		if(mSharedCellStyles == null) {
			mSharedCellStyles = new HashMap<String, CellStyle>();
		}
		mSharedCellStyles.put(styleName, createStyle(cellStyleDefine));
	}
	
	public Color getAwtColor(String color) {
		return getAwtColor(color, Color.black);
	};
	
	public Color getAwtColor(String color, Color defaultColor) {
		if(color.startsWith("#")) {
			return Color.decode(color);
		} else {
			// color is a pre-defined field in {@link Color}
			Color preDefinedColor;
			try {
				Field field = Color.class.getField(color);
				preDefinedColor = (Color) field.get(null);
			} catch (Exception e) {
				e.printStackTrace();
				preDefinedColor = defaultColor;
			}
			return preDefinedColor;
		}
	};
	
	public void autoSizeColumn(int columnIndex) {
		mCurrentSheet.autoSizeColumn(columnIndex);
	}
	
	public void setColumnWidth(int columnIndex, int width) {
		mCurrentSheet.setColumnWidth(columnIndex, width);
	}
	
	public void writeToFile(String filePath) throws IOException {
		FileOutputStream out = new FileOutputStream(filePath);
		mWb.write(out);
		out.close();
		mWb.dispose();
	}
	
	private int getColumnCount(ArrayList<CellDefine> line) {
		int count = 0;
		for(CellDefine cellDefine: line) {
			count += cellDefine.colSpan;
		}
		return count;
	}
	
	private void setBorderStyleAndColor(XSSFCellStyle style, BorderDefine borderDefine, BorderSide side) {
		if(borderDefine != null) {
			BorderStyle borderStyle;
			try {
				borderStyle = BorderStyle.valueOf(borderDefine.style);
				
				if(side.equals(BorderSide.TOP)) {
					style.setBorderTop(borderStyle);
				} else if(side.equals(BorderSide.RIGHT)) {
					style.setBorderRight(borderStyle);
				} else if(side.equals(BorderSide.BOTTOM)) {
					style.setBorderBottom(borderStyle);
				} else {
					style.setBorderLeft(borderStyle);
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
			if(borderDefine.color != null) {
				style.setBorderColor(side, new XSSFColor(getAwtColor(borderDefine.color)));				
			}
		}
	}
	
	private CellStyle createStyle(CellStyleDefine styleDefine) {
		// use XSSFCellStyle because we will generate a xlsx document
		XSSFCellStyle style = (XSSFCellStyle) mWb.createCellStyle();
		// set alignment
		if(styleDefine.alignment != null) {
			try {
				HorizontalAlignment alignment = HorizontalAlignment.valueOf(styleDefine.alignment);
				style.setAlignment(alignment);
			} catch (IllegalArgumentException e) {
				e.printStackTrace();
			}
		}
		if(styleDefine.vertical_alignment != null) {
			try {
				VerticalAlignment verticalAlignment = VerticalAlignment.valueOf(styleDefine.vertical_alignment);
				style.setVerticalAlignment(verticalAlignment);
			} catch (IllegalArgumentException e) {
				e.printStackTrace();
			}
		}
		style.setRotation(styleDefine.rotation);
		style.setWrapText(styleDefine.wrap_text);
		style.setHidden(styleDefine.hidden);
		style.setIndention(styleDefine.indention);
		
		if(styleDefine.border != null ||
				styleDefine.border_top != null ||
				styleDefine.border_right != null || 
				styleDefine.border_bottom != null ||
				styleDefine.border_left != null) {
			// set border styles and color
			BorderDefine[] borderDefines = styleDefine.getBorderDefine();
			setBorderStyleAndColor(style, borderDefines[0], BorderSide.TOP);
			setBorderStyleAndColor(style, borderDefines[1], BorderSide.RIGHT);
			setBorderStyleAndColor(style, borderDefines[2], BorderSide.BOTTOM);
			setBorderStyleAndColor(style, borderDefines[3], BorderSide.LEFT);
			
		}
		
		// set font
		if(styleDefine.font != null) {
			XSSFFont font = (XSSFFont) mWb.createFont();
			if(styleDefine.font.color != null) {
				font.setColor(new XSSFColor(getAwtColor(styleDefine.font.color)));				
			}
			
			font.setItalic(styleDefine.font.italic);
			font.setStrikeout(styleDefine.font.strike_out);
			
			if(styleDefine.font.bold_weight > 0) {
				font.setBoldweight(styleDefine.font.bold_weight);			
			}
			
			if(styleDefine.font.font_family != null) {
				try {
					FontFamily fontFamily = FontFamily.valueOf(styleDefine.font.font_family);
					font.setFamily(fontFamily);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
			
			if(styleDefine.font.font_name != null) {
				font.setFontName(styleDefine.font.font_name);
			}
			
			if(styleDefine.font.underline != null) {
				try {
					FontUnderline underline = FontUnderline.valueOf(styleDefine.font.underline);
					font.setUnderline(underline);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
			
			if(styleDefine.font.height > 0) {
				font.setFontHeight(styleDefine.font.height);
			}
			
			style.setFont(font);
		}
		
		
		// set fill
		
		if(styleDefine.fill != null) {
			if(styleDefine.fill.foreground_color != null) {
				style.setFillForegroundColor(new XSSFColor(getAwtColor(styleDefine.fill.foreground_color)));
			}
			
			if(styleDefine.fill.background_color != null) {
				style.setFillBackgroundColor(new XSSFColor(getAwtColor(styleDefine.fill.background_color)));
			}
			
			if(styleDefine.fill.fill_pattern != null) {
				try {
					FillPatternType fillPattern = FillPatternType.valueOf(styleDefine.fill.fill_pattern);
					style.setFillPattern(fillPattern);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		}
		
		return style;
	}
	
	private void createCell(Row row, int rowIndex, int colIndex, CellDefine cellDefine) {
		Sheet sheet = mCurrentSheet;
		Cell cell = row.createCell(colIndex);
		
		switch(cellDefine.type) {
		case Cell.CELL_TYPE_NUMERIC:
			cell.setCellValue(Double.valueOf(cellDefine.data));
			break;
		case Cell.CELL_TYPE_BLANK:
			cell.setCellValue("");
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			cell.setCellValue(Boolean.valueOf(cellDefine.data));
			break;
		default:
			cell.setCellValue(cellDefine.data);
		}

		cell.setCellType(cellDefine.type);
		
		CellStyle style = null;
		if(cellDefine.customStyle != null) {
			style = createStyle(cellDefine.customStyle);
		} else if(cellDefine.styleName != null) {
			style = mSharedCellStyles.get(cellDefine.styleName);
		}
		
		if(style != null) {
			cell.setCellStyle(style);
		}
		
		// create empty cell for merged regions
		
		if(cellDefine.rowSpan > 1 || cellDefine.colSpan > 1) {
			for(int r = rowIndex; r < rowIndex + cellDefine.rowSpan; r ++) {
				Row rowOfMerged = sheet.getRow(r);
				if(rowOfMerged == null) {
					rowOfMerged = sheet.createRow(r);
				}
				
				for(int c = colIndex; c < colIndex + cellDefine.colSpan; c++) {
					Cell cellOfMerged = rowOfMerged.getCell(c);
					if(cellOfMerged == null) {
						cellOfMerged = rowOfMerged.createCell(c);
					}
					
					if(style != null) {
						cellOfMerged.setCellStyle(style);
					}
				}
			}
		}

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
