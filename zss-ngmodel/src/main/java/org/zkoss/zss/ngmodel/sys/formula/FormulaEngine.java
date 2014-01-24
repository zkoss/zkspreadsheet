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
import org.zkoss.zss.ngmodel.SheetRegion;

/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public interface FormulaEngine {
	
	String KEY_EXTERNAL_BOOK_NAMES = "$ZSS_EXTERNAL_BOOK_NAMES$";

	public FormulaExpression parse(String formula, FormulaParseContext context);
	
	/**
	 * Shift the formula that care about region and formula's absolute, it doesn't care the related sheet.
	 * @param formula
	 * @param srcRegion
	 * @param rowOffset
	 * @param columnOffset
	 * @param context
	 * @return
	 */
	//TODO zss 3.5
//	public FormulaExpression shift(String formula, CellRegion srcRegion, int rowOffset,int columnOffset,
//			FormulaParseContext context);
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
	
	public FormulaExpression renameSheet(String formula, String oldName,String newName,
			FormulaParseContext context);
	
	public EvaluationResult evaluate(FormulaExpression expr, FormulaEvaluationContext context);

	public void clearCache(FormulaClearContext context);
}
