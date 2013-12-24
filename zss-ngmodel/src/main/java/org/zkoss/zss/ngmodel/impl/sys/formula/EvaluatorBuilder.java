/* WorkbookEvaluatorBuilder.java

	Purpose:
		
	Description:
		
	History:
		Dec 24, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
*/
package org.zkoss.zss.ngmodel.impl.sys.formula;

import org.zkoss.poi.ss.formula.EvaluationWorkbook;
import org.zkoss.poi.ss.formula.WorkbookEvaluator;

/**
 * A builder to create workbook evaluator for formula engine 
 * @author Pao
 */
public interface EvaluatorBuilder {
	WorkbookEvaluator createEvaluator(EvaluationWorkbook evalBook);
}
