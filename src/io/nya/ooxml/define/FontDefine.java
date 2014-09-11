package io.nya.ooxml.define;

import org.apache.poi.ss.usermodel.FontFamily;
import org.apache.poi.ss.usermodel.FontUnderline;

public class FontDefine {
	
	public String color;
	
	/**
	 * the enum name of {@link FontFamily}
	 */
	public String font_family;
	
	/**
	 * set the name for the font (i.e. Arial). If the font doesn't exist (because it isn't installed on the system), or the charset is invalid for that font, then another font should be substituted. The string length for this attribute shall be 0 to 31 characters. Default font name is Calibri.
	 */
	public String font_name;
	
	public boolean italic;
	
	public boolean strike_out;
	
	/**
	 * the enum name of {@link FontUnderline}
	 */
	public String underline;
	
	public short bold_weight;
	
	public short height;
	
}
