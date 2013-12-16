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

import org.zkoss.zss.ngmodel.NPrintInfo;

public class PrintInfoImpl implements NPrintInfo,Serializable {
	private static final long serialVersionUID = 1L;
	private boolean printGridline = false; 
	
	@Override
	public boolean isPrintGridline() {
		return printGridline;
	}

	@Override
	public void setPrintGridline(boolean enable) {
		printGridline = enable;
	}

}
