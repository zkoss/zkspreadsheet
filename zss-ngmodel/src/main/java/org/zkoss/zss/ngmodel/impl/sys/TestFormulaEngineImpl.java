package org.zkoss.zss.ngmodel.impl.sys;

import java.io.Serializable;
import java.util.HashMap;
import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.ErrorValue;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult.ResultType;
import org.zkoss.zss.ngmodel.sys.formula.FormulaClearContext;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEvaluationContext;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;
import org.zkoss.zss.ngmodel.sys.formula.FormulaParseContext;
import org.zkoss.zss.ngmodel.util.AreaReference;
import org.zkoss.zss.ngmodel.util.CellReference;
import org.zkoss.zss.ngmodel.util.Validations;

public class TestFormulaEngineImpl implements FormulaEngine {

	
	static HashMap<String, Object[]> testData = new HashMap<String,Object[]>();
	static HashMap<String, Object[]> testType = new HashMap<String,Object[]>();
	
	{
		testData.put("SUM(999)", new Object[]{999D,ResultType.SUCCESS});
		testData.put("SUM(A1)", new Object[]{999D,ResultType.SUCCESS});
		testData.put("SUM(B4:D4)", new Object[]{6D,ResultType.SUCCESS});
		testData.put("A1:A3", new Object[]{new String[]{"A","B","C"},ResultType.SUCCESS});
		testData.put("B1:B3", new Object[]{new Double[]{1D,2D,3D},ResultType.SUCCESS});
		testData.put("C1:C3", new Object[]{new Double[]{4D,5D,6D},ResultType.SUCCESS});
		testData.put("D1", new Object[]{"My Series",ResultType.SUCCESS});
		testData.put("Sheet1!A1:B3", new Object[]{new Double[]{1D,2D,3D},ResultType.SUCCESS});
		testData.put("Sheet2!A$2:B$4", new Object[]{new Double[]{1D,2D,3D},ResultType.SUCCESS});
		
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

		@Override
		public boolean isRefersTo() {
			return true;
		}

		@Override
		public String getRefersToSheetName() {
			AreaReference ref = new AreaReference(formula);
			CellReference cell1 = ref.getFirstCell();
			CellReference cell2 = ref.getLastCell();
			return cell1.getSheetName();
		}

		@Override
		public CellRegion getRefersToCellRegion() {
			AreaReference ref = new AreaReference(formula);
			CellReference cell1 = ref.getFirstCell();
			CellReference cell2 = ref.getLastCell();
			return new CellRegion(cell1.getRow(),cell1.getCol(),cell2.getRow(),cell2.getCol());
		}

		@Override
		public String getRefersToBookName() {
			return null; // TODO zss 3.5
		}

		@Override
		public String getRefersToLastSheetName() {
			return null; // TODO zss 3.5
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

	@Override
	public void clearCache(FormulaClearContext context) {
	}

}
