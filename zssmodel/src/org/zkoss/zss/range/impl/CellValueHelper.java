/* CellValueHelper.java

	Purpose:
		
	Description:
		
	History:
		Jun 24, 2016 5:02:18 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.range.impl;

import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.impl.CellValue;
import org.zkoss.zss.model.impl.CellImpl;
import org.zkoss.zss.model.sys.EngineFactory;
import org.zkoss.zss.model.sys.format.FormatContext;
import org.zkoss.zss.model.sys.format.FormatEngine;

/**
 * @author Henri
 *
 */
public class CellValueHelper {
	FormatEngine _formatEngine;
	
	public static CellValueHelper inst = new CellValueHelper();
	
	public String getFormattedText(SCell cell){
		return getFormatEngine().format(cell, new FormatContext(ZssContext.getCurrent().getLocale())).getText();
	}
	
	public CellValue getCellValue(SCell cell) {
		return ((CellImpl)cell).getEvalCellValue(true);
	}
	
	public Object getValue(SCell cell) {
		return cell.getValue();
	}
	
	protected FormatEngine getFormatEngine(){
		if(_formatEngine==null){
			_formatEngine = EngineFactory.getInstance().createFormatEngine();
		}
		return _formatEngine;
	}
}
