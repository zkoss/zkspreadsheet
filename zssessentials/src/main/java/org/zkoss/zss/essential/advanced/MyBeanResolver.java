package org.zkoss.zss.essential.advanced;

import org.zkoss.xel.VariableResolver;
import org.zkoss.xel.XelException;

public class MyBeanResolver implements VariableResolver {

	private static AssetsBean assetsBean = new AssetsBean();
	
	@Override
	public Object resolveVariable(String name) throws XelException {
		if (name.equals("assetsBean")){
			// find the bean in your environment
			return assetsBean;
		}
		return null;
	}
}
