/* DependencyTableAdv.java

	Purpose:
		
	Description:
		
	History:
		Dec 12, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.ngmodel.impl.sys;

import java.io.Serializable;
import org.zkoss.zss.ngmodel.NBookSeries;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyTable;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;

/**
 * @author Pao
 * @since 3.5
 */
public abstract class DependencyTableAdv implements DependencyTable, Serializable {
	private static final long serialVersionUID = 1L;
	
	abstract public void setBookSeries(NBookSeries series);

	abstract public void add(Ref dependant, Ref precedent);

	abstract public void merge(DependencyTableAdv dependencyTable);
}
