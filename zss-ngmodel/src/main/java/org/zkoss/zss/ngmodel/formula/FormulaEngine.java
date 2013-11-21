package org.zkoss.zss.ngmodel.formula;

public interface FormulaEngine {

	FormulaExpression parse(String formula, FormulaParseContext context);
	
	Object evaluate(FormulaExpression expr);
}
