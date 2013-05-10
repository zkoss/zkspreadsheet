package org.zkoss.zss.api.model;


public interface Font {
	
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
	
	
	public Color getColor();
	public String getFontName();
	public Boldweight getBoldweight();
	public short getFontHeight();
	public boolean isItalic();
	public boolean isStrikeout();
	public TypeOffset getTypeOffset();
	public Underline getUnderline();
	public void setFontName(String fontName);
	public void setBoldweight(Boldweight boldweight);
	public void setItalic(boolean italic);
	public void setStrikeout(boolean strikeout);
	public void setUnderline(Underline underline);
	public void setFontHeight(short height);
	public void setColor(Color color);
	
}
