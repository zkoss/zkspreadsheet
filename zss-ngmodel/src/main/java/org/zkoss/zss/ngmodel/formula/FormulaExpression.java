package org.zkoss.zss.ngmodel.formula;

public interface FormulaExpression {

	String getFormulaString();
	
	String reformSheetNameChanged(String oldName,String newName);
}
