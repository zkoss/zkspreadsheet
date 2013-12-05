package org.zkoss.zss.ngmodel.impl;

public interface LinkedModelObject {
	/**
	 * Destroy / release this model object, for example all the dependency, parent linking.
	 * this method has to be called before remove this linking from parent object 
	 */
	public void destroy();
	
	public void checkOrphan();
}
