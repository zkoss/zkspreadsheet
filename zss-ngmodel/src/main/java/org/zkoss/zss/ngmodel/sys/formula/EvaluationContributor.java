package org.zkoss.zss.ngmodel.sys.formula;

import org.zkoss.xel.FunctionMapper;
import org.zkoss.xel.VariableResolver;
import org.zkoss.zss.ngmodel.NBook;

public interface EvaluationContributor {

	public FunctionMapper getFunctionMaper(NBook book);
	
	public VariableResolver getVariableResolver(NBook book);
}
