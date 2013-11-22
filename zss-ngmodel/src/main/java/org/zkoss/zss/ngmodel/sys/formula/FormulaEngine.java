package org.zkoss.zss.ngmodel.sys.formula;

public interface FormulaEngine {

	FormulaExpression parse(String formula, FormulaParseContext context);
	
	EvaluationResult evaluate(FormulaExpression expr, FormulaEvaluationContext context);
}
