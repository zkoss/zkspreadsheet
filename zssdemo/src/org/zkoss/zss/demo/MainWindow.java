/* MainWindow.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Dec 20, 2007 12:40:46 PM     2007, Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.demo;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.ext.AfterCompose;
//import org.zkoss.zss.engine.xel.Indexes;
import org.zkoss.zss.model.Book;
/*import org.zkoss.zss.model.Cell;
import org.zkoss.zss.model.Sheet;
*/import org.zkoss.zss.ui.event.CellEvent;
import org.zkoss.zss.ui.event.Events;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Label;
import org.zkoss.zul.Textbox;
import org.zkoss.zul.Window;

/**
 * @author Dennis.Chen
 *
 */
public class MainWindow extends Window implements AfterCompose{

	private static final long serialVersionUID = 1;
	
	int lastRow=0;
	int lastCol=0;
	Book book;
	Spreadsheet spreadsheet;
	
	
	public void afterCompose() {
		spreadsheet = (Spreadsheet)getFellow("ss1");
		book = spreadsheet.getBook();
		
		spreadsheet.addEventListener(Events.ON_CELL_FOUCSED,new EventListener(){
			public void onEvent(Event event) throws Exception {
				onCellEvent((CellEvent)event);
			}
		});
		spreadsheet.addEventListener(Events.ON_START_EDITING,new EventListener(){
			public void onEvent(Event event) throws Exception {
				onCellEvent((CellEvent)event);
			}
		});
		
		final Textbox tbxval = (Textbox)getFellow("tbxval");
		tbxval.addEventListener("onChange",new EventListener(){
			public void onEvent(Event event) throws Exception {
				doCellChange(tbxval.getValue());
			}
		});
	}
	void onCellEvent(CellEvent event){
		Sheet sheet = event.getSheet();
		lastRow = event.getRow();
		lastCol = event.getColumn();
		Label lbpos = (Label)getFellow("lbpos");
		Textbox tbxval = (Textbox)getFellow("tbxval");
		CellRangeAddress addr = new CellRangeAddress(lastRow, lastRow, lastCol, lastCol);
		Cell cell = Utils.getCell(sheet, lastRow, lastCol);
		lbpos.setValue(addr.formatAsString());
		tbxval.setValue(cell == null ? "" : Utils.getEditText(cell));
	}
		
	void doCellChange(String value){
		if(lastRow == -1){
			return;
		}
		Sheet sheet = (Sheet)book.getSheetAt(0);
		Cell cell = Utils.getOrCreateCell(sheet, lastRow, lastCol);
		Utils.setEditText(cell, value);
	}

}
