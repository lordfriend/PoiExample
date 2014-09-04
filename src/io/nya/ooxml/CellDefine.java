package io.nya.ooxml;

public class CellDefine {

	// raw data
	public String data;
	
	/**
	 *  cell type of {@link Cell}
	 *  String type is 1;
	 *  numeric type is 0;
	 *  formula type is 2;
	 *  error type is 5;
	 *  boolean type is 4;
	 *  blank type is 3;
	 */
	public int type;
	
	// cell merges
	public int rowSpan = 1;
	public int colSpan = 1;
	
	// style name predefined in style section. when customStyle is defined. this field is ignored.
	public String styleName;
	
	// custom style for this cell
	public CellStyleDefine customStyle;
}
