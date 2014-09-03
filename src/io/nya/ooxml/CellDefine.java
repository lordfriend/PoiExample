package io.nya.ooxml;

public class CellDefine {

	// raw data
	public String data;
	
	// data type
	public String type;
	
	// cell merges
	public int rowSpan = 1;
	public int colSpan = 1;
	
	// style name predefined in style section. when customStyle is defined. this field is ignored.
	public String styleName;
	
	// custom style for this cell
	public CellStyle customStyle;
}
