/* DesktopCellStyleContext.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 15, 2010 9:27:22 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul.ctrl;

import org.zkoss.zk.ui.Desktop;
import org.zkoss.zss.app.Consts;

/**
 * @author Sam
 *
 */
public class DesktopCellStyleContext extends AbstractBaseContext implements CellStyleContext {
	
	CellStyleApplier cellStyle;
	
	public void doTargetChange(CellStyleApplier aFontStyle){
		cellStyle = aFontStyle;
		
		CellStyleContextEvent event = new CellStyleContextEvent(
				Consts.ON_STYLING_TARGET_CHANGED, cellStyle);
		
		listenerStore.fire(event);
	}

	public CellStyleApplier getCellStyle(){
		return cellStyle;
	}


	@Override
	public void modifyStyle(StyleModification styleModification) {
		
		CellStyleContextEvent candidteEvt = new CellStyleContextEvent(
				Consts.ON_CELL_STYLE_CHANGED, getCellStyle());
		
		styleModification.modify(cellStyle, candidteEvt);
		
		listenerStore.fire(candidteEvt);
	}
	
}