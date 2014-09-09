package io.nya.ooxml;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public class CellStyleDefine {
	
	public BorderDefine border;
	
	/**
	 * Side specified style will override the default
	 */
	public BorderDefine border_top;
	public BorderDefine border_right;
	public BorderDefine border_bottom;
	public BorderDefine border_left;
	
	public FontDefine font;
	
	public FillDefine fill;
	
	/**
	 * the enum name of {@link HorizontalAlignment}
	 */
	public String alignment;
	
	/**
	 * the enum name of {@link VerticalAlignment}
	 */
	public String vertical_alignment;
	
	public short rotation;
	
	public boolean wrap_text;
	
	public short indention;
	
	public boolean hidden;
	
	/**
	 * Get the border define array with the order of top, right, bottom, left.
	 * The array is the merged result of border and the correspond side border
	 * @return
	 */
	public BorderDefine[] getBorderDefine() {
		BorderDefine[] defines = new BorderDefine[4];
		defines[0] = border_top != null ? border_top : border;
		defines[1] = border_right != null ? border_right : border;
		defines[2] = border_bottom != null ? border_bottom : border;
		defines[3] = border_left != null ? border_left : border;
		return defines;
	}
}
