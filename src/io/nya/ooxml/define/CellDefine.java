package io.nya.ooxml.define;

public class CellDefine {
	
	public CellDefine() {
		// default constructor
	}
	
	public CellDefine(String data) {
		this.data = data;
	}
	
	public CellDefine(String data, String styleName) {
		this.data = data;
		this.styleName = styleName;
	}
	
	public CellDefine(String data, int type, String styleName) {
		this.data = data;
		this.type = type;
		this.styleName = styleName;
	}

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
	public int type = 1;
	
	// cell merges
	public int rowSpan = 1;
	public int colSpan = 1;
	
	// style name predefined in style section. when customStyle is defined. this field is ignored.
	public String styleName;
	
	public short height = -1;
	
	// custom style for this cell
	public CellStyleDefine customStyle;
}
