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
package org.zkoss.zss.model.sys.formula;

import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SheetRegion;

/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public interface FormulaEngine {
	
	String KEY_EXTERNAL_BOOK_NAMES = "$ZSS_EXTERNAL_BOOK_NAMES$";

	public FormulaExpression parse(String formula, FormulaParseContext context);
	
	/**
	 * Shift the formula base on the offset
	 * @param formula
	 * @param rowOffset
	 * @param columnOffset
	 * @param context
	 * @return
	 */
	public FormulaExpression shift(String formula, int rowOffset,int columnOffset, FormulaParseContext context);
	
	/**
	 * Transpose the formula base one the origin
	 * @param formula
	 * @param rowOrigin
	 * @param columnOrigin
	 * @param context
	 * @return
	 */
	public FormulaExpression transpose(String formula, int rowOrigin,int columnOrigin, FormulaParseContext context);
	/**
	 * Shift the formula that care on sheet and region.
	 * @param formula
	 * @param srcRegion
	 * @param rowOffset
	 * @param columnOffset
	 * @param context
	 * @return
	 */
	public FormulaExpression move(String formula, SheetRegion srcRegion, int rowOffset,int columnOffset,
			FormulaParseContext context);
	
	public FormulaExpression shrink(String formula, SheetRegion srcRegion, boolean hrizontal,
			FormulaParseContext context);
	
	public FormulaExpression extend(String formula, SheetRegion srcRegion, boolean hrizontal,
			FormulaParseContext context);
	
	public FormulaExpression renameSheet(String formula, SBook book, String oldName,String newName,
			FormulaParseContext context);
	
	public EvaluationResult evaluate(FormulaExpression expr, FormulaEvaluationContext context);

	public void clearCache(FormulaClearContext context);
}
