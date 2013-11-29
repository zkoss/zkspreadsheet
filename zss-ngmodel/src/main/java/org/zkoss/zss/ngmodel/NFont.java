package org.zkoss.zss.ngmodel;

public interface NFont {
	public enum TypeOffset{
		NONE, 
		SUPER, 
		SUB
	}
	public enum Underline{
		NONE,
		SINGLE,
		DOUBLE,
		SINGLE_ACCOUNTING,
		DOUBLE_ACCOUNTING
	}
	
	public enum Boldweight{
		NORMAL,
		BOLD
	}
	
	/**
	 * 
	 * @return a font's color
	 */
	public String getColor();
	
	/**
	 * 
	 * @return a font's name like "Calibri".
	 */
	public String getName();
	
	/**
	 * 
	 * @return a font's bold style.
	 */
	public Boldweight getBoldweight();
	
	/**
	 * 
	 * @return a fon't height in pixel
	 */
	public int getHeight();
	
	/**
	 * 
	 * @return true if the font is italic
	 */
	public boolean isItalic();
	
	/**
	 * 
	 * @return true if the font is strike-out.
	 */
	public boolean isStrikeout();
	
	/**
	 * 
	 * @return
	 */
	public TypeOffset getTypeOffset();
	
	/**
	 * 
	 * @return the style of a font's underline
	 */
	public Underline getUnderline();
	
	
	public void setName(String fontName);

	public void setColor(String fontColor);

	public void setBoldweight(Boldweight fontBoldweight);

	public void setHeight(int fontHeight);

	public void setItalic(boolean fontItalic);

	public void setStrikeout(boolean fontStrikeout);

	public void setTypeOffset(TypeOffset fontTypeOffset);

	public void setUnderline(Underline fontUnderline);
}
