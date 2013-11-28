package org.zkoss.zss.ngmodel.impl;

public interface LinkedModelObject {
	/**
	 * Release this model object, for example all the dependency, parent linking.
	 * this method has to be called before remove this linking from parent object 
	 */
	void release();
	void checkOrphan();
}
