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

import java.util.List;

/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public interface NDataValidation extends FormulaContent{
	public enum ErrorStyle {
		STOP((byte)0x00), WARNING((byte)0x01), INFO((byte)0x02);
		
		private byte value;
		ErrorStyle(byte value){
			this.value = value;
		}
		
		public byte getValue(){
			return value;
		}
	}
	
	public enum ValidationType {
		ANY, INTEGER, DECIMAL, LIST, 
		DATE, TIME, TEXT_LENGTH, FORMULA;

	}
	public enum OperatorType{
		BETWEEN, NOT_BETWEEN, EQUAL, NOT_EQUAL, 
		GREATER_THAN, LESS_THAN, GREATER_OR_EQUAL, LESS_OR_EQUAL;
	}
	
	public NSheet getSheet();
	
	public ErrorStyle getErrorStyle();
	public void setErrorStyle(ErrorStyle errorStyle);
	
	public void setEmptyCellAllowed(boolean allowed);
	public boolean isEmptyCellAllowed();
	
	public void setShowDropDownArrow(boolean show);
	public boolean isShowDropDownArrow();
	
	public void setShowPromptBox(boolean show);
	public boolean isShowPromptBox();

	public void setShowErrorBox(boolean show);
	public boolean isShowErrorBox();

	public void setPromptBox(String title, String text);
	public String getPromptBoxTitle();
	public String getPromptBoxText();

	public void setErrorBox(String title, String text);
	public String setErrorBoxTitle();
	public String getErrorBoxText();

	public List<CellRegion> getRegions();
	public void addRegion(CellRegion region);
	public void removeRegion(CellRegion region);
	
	
	public ValidationType getValidationType();
	public void setValidationType(ValidationType type);
	
	public OperatorType getOperatorType();
	public void setOperatorType(OperatorType type);
	
	/**
	 * Return formula parsing state.
	 * @return true if has error, false if no error or no formula
	 */
	public boolean isFormulaParsingError();
	
	public int getNumOfValue1();
	public Object getValue1(int i);
	public int getNumOfValue2();
	public Object getValue2(int i);
	
	public String getValue1Formula();
	public String getValue2Formula();
	public void setFormula(String valueExpr);
	public void setFormula(String value1Expr,String value2Expr);

	public Object getId();
}
