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
