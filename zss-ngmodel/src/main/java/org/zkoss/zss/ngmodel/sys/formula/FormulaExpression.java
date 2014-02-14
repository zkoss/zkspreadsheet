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
package org.zkoss.zss.ngmodel.sys.formula;

import org.zkoss.zss.ngmodel.CellRegion;

/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public interface FormulaExpression {

	/**
	 * indicated the expression has parsing error
	 * @return
	 */
	boolean hasError();
	
	/**
	 * Get the expression parsing error message if any
	 * @return
	 */
	String getErrorMessage();
	
	String getFormulaString();
	
//	String reformSheetNameChanged(String oldName,String newName);
	
	//parsing result for Name
	
	boolean isRefersTo();

	String getRefersToBookName();
	
	String getRefersToSheetName();
	String getRefersToLastSheetName();
	
	CellRegion getRefersToCellRegion();
}
