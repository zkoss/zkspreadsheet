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

import org.zkoss.zss.ngmodel.NViewInfo;

public class ViewInfoImpl implements NViewInfo,Serializable {

	
	private boolean displayGridline = true; 
	
	
	@Override
	public boolean isDisplayGridline() {
		return displayGridline;
	}

	@Override
	public void setDisplayDridline(boolean display) {
		displayGridline = display;
	}

}
