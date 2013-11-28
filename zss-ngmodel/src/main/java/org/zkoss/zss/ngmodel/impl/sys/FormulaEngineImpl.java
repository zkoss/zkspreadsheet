package org.zkoss.zss.ngmodel.impl.sys;

import java.io.Serializable;
import java.util.HashMap;

import org.zkoss.zss.ngmodel.ErrorValue;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult.ResultType;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEvaluationContext;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;
import org.zkoss.zss.ngmodel.sys.formula.FormulaParseContext;
import org.zkoss.zss.ngmodel.util.Validations;

public class FormulaEngineImpl implements FormulaEngine {

	
	static HashMap<String, Object[]> testData = new HashMap<String,Object[]>();
	static HashMap<String, Object[]> testType = new HashMap<String,Object[]>();
	
	{
		testData.put("SUM(999)", new Object[]{999,ResultType.SUCCESS});
		testData.put("SUM(A1)", new Object[]{999,ResultType.SUCCESS});
		testData.put("A1:A3", new Object[]{new String[]{"A","B","C"},ResultType.SUCCESS});
		testData.put("B1:B3", new Object[]{new Integer[]{1,2,3},ResultType.SUCCESS});
		testData.put("C1:C3", new Object[]{new Integer[]{4,5,6},ResultType.SUCCESS});
	}
	
	
	public FormulaExpression parse(final String formula, FormulaParseContext context) {
		Validations.argNotNull(formula,context);
		return new FormulaExpressionImpl(formula);
	}
	
	static class FormulaExpressionImpl implements FormulaExpression,Serializable{
		String formula;
		public FormulaExpressionImpl(String formula){
			this.formula = formula;
		}
		
		public String getFormulaString() {
			return formula;
		}

		public String reformSheetNameChanged(String oldName, String newName) {
			return formula;
		}

		public boolean hasError() {
			return testData.get(formula)==null;
		}
	}

	public EvaluationResult evaluate(final FormulaExpression expr, FormulaEvaluationContext context) {
		Validations.argNotNull(expr,context);
		return new EvaluationResult(){

			public ResultType getType() {
				return expr.hasError()?ResultType.ERROR:(ResultType)testData.get(expr.getFormulaString())[1];
			}

			public Object getValue() {
				return expr.hasError()?new ErrorValue(ErrorValue.INVALID_FORMULA,"formula parsing error"):testData.get(expr.getFormulaString())[0];
			}};
	}

}
