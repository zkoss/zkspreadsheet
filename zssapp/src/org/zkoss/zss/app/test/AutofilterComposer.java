/* AutofilterComposer.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Feb 29, 2012 9:10:05 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.app.test;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * @author sam
 *
 */
public class AutofilterComposer extends GenericForwardComposer {
	
	private Spreadsheet spreadsheet;
	private Worksheet sheet;
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		sheet = spreadsheet.getSelectedSheet();
		initAutofilter();
	}

	
	private void prepareDummyData() {
		Range rng = Ranges.range(sheet);
		int colSize = 10;
		int rowSize = 50;
		
		for (int c = 0; c < colSize; c++) {
			for (int r = 0; r < rowSize; r++) {
				Ranges.range(sheet, r, c).setEditText(c + "_" + r);
			}
		}
	}
	
	private void initAutofilter() {
		prepareDummyData();
		
		Ranges.range(sheet, 0, 0).autoFilter();
	}
	
	
	
}
