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

import org.zkoss.zss.ngmodel.SheetRegion;

/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public interface FormulaEngine {
	
	String KEY_EXTERNAL_BOOK_NAMES = "$ZSS_EXTERNAL_BOOK_NAMES$";

	public FormulaExpression parse(String formula, FormulaParseContext context);
	
	public FormulaExpression shift(String formula, SheetRegion srcRegion, int rowOffset,int columnOffset,
			FormulaParseContext context);
	
	public FormulaExpression renameSheet(String formula, String oldName,String newName,
			FormulaParseContext context);
	
	public EvaluationResult evaluate(FormulaExpression expr, FormulaEvaluationContext context);

	public void clearCache(FormulaClearContext context);
}
