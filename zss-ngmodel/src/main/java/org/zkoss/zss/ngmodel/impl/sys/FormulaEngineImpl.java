package org.zkoss.zss.ngmodel.impl.sys;

import org.zkoss.zss.ngmodel.ErrorValue;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEvaluationContext;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;
import org.zkoss.zss.ngmodel.sys.formula.FormulaParseContext;
import org.zkoss.zss.ngmodel.util.Validations;

public class FormulaEngineImpl implements FormulaEngine {

	public FormulaExpression parse(final String formula, FormulaParseContext context) {
		Validations.argNotNull(formula,context);
		return new FormulaExpression(){

			public String getFormulaString() {
				return formula;
			}

			public String reformSheetNameChanged(String oldName, String newName) {
				return formula;
			}

			public boolean hasError() {
				return !formula.startsWith("SUM(999)");
			}};
	}

	public EvaluationResult evaluate(final FormulaExpression expr, FormulaEvaluationContext context) {
		Validations.argNotNull(expr,context);
		return new EvaluationResult(){

			public CellType getType() {
				return expr.hasError()?CellType.ERROR:CellType.NUMBER;
			}

			public Object getValue() {
				return expr.hasError()?new ErrorValue((byte)0,"formula parsing error"):999;
			}};
	}

}
