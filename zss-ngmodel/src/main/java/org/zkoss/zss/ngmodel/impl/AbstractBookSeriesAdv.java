/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;

import org.zkoss.zss.ngmodel.NBookSeries;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyTable;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public abstract class AbstractBookSeriesAdv implements NBookSeries,Serializable{
	private static final long serialVersionUID = 1L;

	protected boolean autoFormulaCacheClean = true;//default on
	
	public abstract DependencyTable getDependencyTable();

	@Override
	public boolean isAutoFormulaCacheClean() {
		return autoFormulaCacheClean;
	}
	@Override
	public void setAutoFormulaCacheClean(boolean autoFormulaCacheClean) {
		this.autoFormulaCacheClean = autoFormulaCacheClean;
	}
	
	
	
}
