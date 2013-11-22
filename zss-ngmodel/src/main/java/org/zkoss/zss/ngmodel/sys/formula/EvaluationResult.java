package org.zkoss.zss.ngmodel.sys.formula;

import org.zkoss.zss.ngmodel.NCell.CellType;

public interface EvaluationResult {

	CellType getType();
	Object getValue();
	
}
