package org.zkoss.zss.ngmodel;

import java.util.Date;

import org.zkoss.zss.ngmodel.impl.SheetAdv;

public interface NCell {

	public enum CellType {
		BLANK,
		STRING,
		RICHTEXT,
		FORMULA,
		NUMBER,
		BOOLEAN,		
		DATE,
		ERROR
	}
	
	public NSheet getSheet();
	
	public CellType getType();
	public CellType getFormulaResultType();
	
	public boolean isNull();
	
	public int getRowIndex();
	
	public int getColumnIndex();
	
	public String getReferenceString();
	
	public NCellStyle getCellStyle();
	
	public NCellStyle getCellStyle(boolean local);
	
	public void setCellStyle(NCellStyle cellStyle);
	
	public NHyperlink getHyperlink();
	
	/**
	 * Set or clear a hyperlink
	 * @param hyperlink hyperlink to set, or null to clear
	 */
	public void setHyperlink(NHyperlink hyperlink);
	
	/** set a empty hyperlinkt*/
	public NHyperlink setHyperlink();

//	boolean isReadonly();
//	
	
	public Object getValue();
	public void setValue(Object value);

	//clear cell value , reset it to blank
	public void clearValue();//
	public void clearFormulaResultCache();
	
	public void setStringValue(String value);
	public String getStringValue();
	
	public void setRichTextValue(NRichText text);
	
	/** set a empty rich text **/
	public NRichText setRichTextValue();
	
	public NRichText getRichTextValue();
	
	/**
	 * set formula as string with '=', ex: SUM(A1:B2)
	 * @param fromula
	 */
	public void setFormulaValue(String formula);
	public String getFormulaValue();
	
	public void setNumberValue(Number number);
	public Number getNumberValue();
	
	public void setDateValue(Date date);
	public Date getDateValue();
	
	public void setBooleanValue(Boolean bool);
	public Boolean getBooleanValue();
	
	public ErrorValue getErrorValue();
	public void setErrorValue(ErrorValue errorValue);
	
	
	public void setComment(NComment comment);
	public NComment setComment();
	public NComment getComment();
	
}
