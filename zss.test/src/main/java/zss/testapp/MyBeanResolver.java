package zss.testapp;

import org.zkoss.xel.VariableResolver;
import org.zkoss.xel.XelException;

/**
 * A example of VariableResolver to get JavaBeans
 * @author Hawk
 *
 */
public class MyBeanResolver implements VariableResolver {
	
	@Override
	public Object resolveVariable(String name) throws XelException {
		return MyBeanService.getMyBeanService().get(name);
	}
}
