package org.zkoss.zss.ngmodel.sys.formula;

public interface FormulaExpression {

	/**
	 * indicated the expression has parsing error
	 * @return
	 */
	boolean hasError();
	
	String getFormulaString();
	
	String reformSheetNameChanged(String oldName,String newName);
}
