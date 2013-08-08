package org.zkoss.zss.essential.advanced;

import org.zkoss.xel.VariableResolver;
import org.zkoss.xel.XelException;

public class MyBeanResolver implements VariableResolver {

	private static MyBeanService myBeanService = MyBeanService.getMyBeanService();
	
	@Override
	public Object resolveVariable(String name) throws XelException {
		return myBeanService.get(name);
	}
}
