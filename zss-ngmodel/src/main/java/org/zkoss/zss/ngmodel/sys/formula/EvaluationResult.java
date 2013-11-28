package org.zkoss.zss.ngmodel.sys.formula;

import org.zkoss.zss.ngmodel.NCell.CellType;

public interface EvaluationResult {

	public enum ResultType{
		SUCCESS,ERROR
	}
	
	ResultType getType();
	Object getValue();
	
}
