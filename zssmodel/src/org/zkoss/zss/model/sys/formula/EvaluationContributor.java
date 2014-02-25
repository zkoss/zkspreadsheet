package org.zkoss.zss.model.sys.formula;

import org.zkoss.xel.FunctionMapper;
import org.zkoss.xel.VariableResolver;
import org.zkoss.zss.model.SBook;

public interface EvaluationContributor {

	public FunctionMapper getFunctionMaper(SBook book);
	
	public VariableResolver getVariableResolver(SBook book);
}
