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
	private static final long serialVersionUID = 1L;
	private boolean displayGridline = true;
	
	private int rowFreeze = 0;
	private int columnFreeze = 0;
	
	
	
	@Override
	public boolean isDisplayGridline() {
		return displayGridline;
	}

	@Override
	public void setDisplayGridline(boolean enable) {
		displayGridline = enable;
	}

	@Override
	public int getNumOfRowFreeze() {
		return rowFreeze;
	}

	@Override
	public int getNumOfColumnFreeze() {
		return columnFreeze;
	}

	@Override
	public void setNumOfRowFreeze(int num) {
		rowFreeze = num;
	}

	@Override
	public void setNumOfColumnFreeze(int num) {
		columnFreeze = num;
	}
}
