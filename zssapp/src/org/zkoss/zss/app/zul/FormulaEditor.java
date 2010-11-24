/* FormulaEditor.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 15, 2010 4:28:32 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.InputEvent;
import org.zkoss.zss.ui.Position;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.CellEvent;
import org.zkoss.zss.ui.event.EditboxEditingEvent;
import org.zkoss.zss.ui.event.Events;
import org.zkoss.zss.ui.event.StopEditingEvent;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zul.Textbox;

/**
 * @author Sam
 *
 */
public class FormulaEditor extends Textbox implements ZssappComponent{

	private Spreadsheet ss;
	
	private Cell currentEditcell;
	
	public FormulaEditor() {
		
	}
	
	public void onChanging(InputEvent event) {
		if (currentEditcell == null)
			return;
		if (currentEditcell.getCellType() == Cell.CELL_TYPE_FORMULA
				&& currentEditcell.getCellFormula() != null)
			return;

		Utils.setEditText(currentEditcell, ((InputEvent) event).getValue());
	}
	
	public void onFocus() {
		
		int left = ss.getSelection().getLeft();
		int top = ss.getSelection().getTop();
		Sheet sheet = ss.getSelectedSheet();
		currentEditcell = Utils.getCell(sheet, top, left);
	}
	
	public void onBlur() {
		currentEditcell = null;
	}
	
	public void onOK() {
		Utils.setEditText(ss.getSelectedSheet(), ss
				.getSelection().getTop(), ss.getSelection().getLeft(),
				getText());
		Position pos = ss.getCellFocus();
		int row = pos.getRow() + 1;
		int col = pos.getColumn();

		ss.focusTo(row, col);
	}

	@Override
	public void bindSpreadsheet(Spreadsheet spreadsheet) {
		ss = spreadsheet;
		
		setWidgetListener("onFocus", "this.$f('" + ss.getId() + "', true).focus(false);");
		
		ss.addEventListener(Events.ON_CELL_FOUCSED, new EventListener() {

			@Override
			public void onEvent(Event event) throws Exception {
				// TODO Auto-generated method stub
				CellEvent evt = (CellEvent)event;
				int row = evt.getRow();
				int col = evt.getColumn();
				Cell cell = Utils.getCell(ss.getSelectedSheet(), row, col);
				String editText = Utils.getEditText(cell);
				setText(cell == null ? "" : (editText == null ? "" : editText));
			}
			
		});
		ss.addEventListener(Events.ON_STOP_EDITING,
				new EventListener() {
					public void onEvent(Event event) throws Exception {
						StopEditingEvent evt = (StopEditingEvent)event;
						
						setText((String) evt.getEditingValue());

						//chart not implement yet
//						// to notify all widgets there is a cell changed
//						for (int i = 0; i < chartKey; i++) {
//							try {
//								Window win = (Window) mainWin.getFellow("chartWin" + i);
//								if (win != null) {
//									Chart myChart = (Chart) win.getFellow("myChart");
//									CellEvent event = new StopEditingEvent(
//											org.zkoss.zss.ui.event.Events.ON_STOP_EDITING,
//											myChart, evt.getSheet(), evt.getRow(), evt
//													.getColumn(), (String) evt.getData());
//									org.zkoss.zk.ui.event.Events.postEvent(event);
//								}
//							} catch (Exception e) {
//								// the chart may be deleted
//							}
//						}
					}

				});
		ss.addEventListener(Events.ON_EDITBOX_EDITING,
				new EventListener() {
			public void onEvent(Event event) throws Exception {
				EditboxEditingEvent evt = (EditboxEditingEvent)event;
				setText((String) evt.getEditingValue());
			}
		});
	}

	@Override
	public void unbindSpreadsheet() {
		// TODO Auto-generated method stub
		
	}

}
