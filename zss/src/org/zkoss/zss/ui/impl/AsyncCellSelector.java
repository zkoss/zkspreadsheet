/* AsyncCellSelector.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Sep 17, 2010 5:47:24 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 3.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.ui.impl;

import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.ui.Rect;

/**
 * @author Sam
 *
 */
public class AsyncCellSelector extends CellSelector {
	
	public void doVisit(final XSheet sheet, final Rect rect, final CellVisitor vistor){
		new Thread(){
			public void run() {
				AsyncCellSelector.super.doVisit(sheet, rect, vistor);
			}
		}.start();
		
	
	}
}
