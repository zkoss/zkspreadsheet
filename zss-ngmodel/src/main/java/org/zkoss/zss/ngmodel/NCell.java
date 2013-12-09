/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngmodel;

import java.util.Date;

import org.zkoss.zss.ngmodel.impl.SheetAdv;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public interface NCell extends FormulaContent{

	public enum CellType {
		BLANK(3),
		STRING(1),
		FORMULA(2),
		NUMBER(0),
		BOOLEAN(4),		
		ERROR(5);
		
		private int value;
	    private CellType(int value) {
	        this.value = value;
	    }

	    public int value() {
	        return value;
	    }
	    
//	    public String toString(){
//	    	new Exception().printStackTrace();
//	    	return "XXX";
//	    }
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
	
	public boolean isDefaultCellStyle();
	
	public void setCellStyle(NCellStyle cellStyle);
	
	public NHyperlink getHyperlink();
	
	/**
	 * Set or clear a hyperlink
	 * @param hyperlink hyperlink to set, or null to clear
	 */
	public void setHyperlink(NHyperlink hyperlink);
	
	/** set a empty hyperlinkt*/
	public NHyperlink setupHyperlink();

//	boolean isReadonly();
//	
	
	public Object getValue();
	public void setValue(Object value);

	/**
	 * clear cell value , reset it to blank
	 */
	public void clearValue();//
	/**
	 * Set string value, if the value start with '=', then it sets as formula 
	 */
	public void setStringValue(String value);
	public String getStringValue();
	
	/**
	 * Set rich test value.
	 */
	public void setRichText(NRichText text);
	
	/** Set a empty rich text value and return the instance which to be edited **/
	public NRichText setupRichText();
	
	public NRichText getRichText();
	
	/**
	 * set formula as string without '=', ex: SUM(A1:B2)
	 * @param fromula
	 */
	public void setFormulaValue(String formula);
	public String getFormulaValue();
	
	public void setNumberValue(Number number);
	public Number getNumberValue();
	
	/**
	 * Sets the number value a date instance, it will transfer the date to double value 
	 */
	public void setDateValue(Date date);
	/**
	 * Gets the date value that is transfered by the double number value.
	 */
	public Date getDateValue();
	
	public void setBooleanValue(Boolean bool);
	/**
	 * Gets the boolean value
	 */
	public Boolean getBooleanValue();
	
	public ErrorValue getErrorValue();
	public void setErrorValue(ErrorValue errorValue);
	
	
	public void setComment(NComment comment);
	public NComment setupComment();
	public NComment getComment();
	
}
