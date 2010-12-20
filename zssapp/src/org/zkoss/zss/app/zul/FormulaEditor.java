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
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.InputEvent;
import org.zkoss.zss.app.cell.EditHelper;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.impl.SheetCtrl;
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
	
	private String initVal;
	
	private boolean focusOut = false;
	private Cell currentEditcell;
	
	public FormulaEditor() {
		setCols(100);
	}
	
	public void onChanging(InputEvent event) {
		if (currentEditcell == null)
			return;
		if (currentEditcell.getCellType() == Cell.CELL_TYPE_FORMULA
				&& currentEditcell.getCellFormula() != null)
			return;

		Utils.setEditText(currentEditcell, ((InputEvent) event).getValue());
	}

	public void onCancel() {
		System.out.println("onCancel");
		focusOut = true;
		int row = ss.getSelection().getTop();
		int col = ss.getSelection().getLeft();
		if (initVal != null) {
			Utils.setEditText(ss.getSelectedSheet(), row, col, initVal);
			initVal = null;
		}
		ss.focusTo(row, col);
		
	}
	
	public void onFocus() {
		focusOut = false;
		int left = ss.getSelection().getLeft();
		int top = ss.getSelection().getTop();
		Worksheet sheet = ss.getSelectedSheet();
		currentEditcell = Utils.getCell(sheet, top, left);
		
		if (currentEditcell != null)
			initVal = Ranges.range(sheet, top, left).getEditText();
		EditHelper.clearCutOrCopy(ss);
	}
	
	public void onBlur() {
		if (focusOut) {
			focusOut = true;
			return;
		}
		initVal = null;
		
		Utils.setEditText(ss.getSelectedSheet(), ss
				.getSelection().getTop(), ss.getSelection().getLeft(),
				getText());
		Position pos = ss.getCellFocus();
		int row = pos.getRow();
		int col = pos.getColumn();
		final Worksheet sheet = ss.getSelectedSheet();
		final CellRangeAddress merged = sheet != null ? ((SheetCtrl)sheet).getMerged(row, col) : null;
		col = merged == null ? col + 1 : merged.getLastColumn() + 1;
		if (ss.getMaxcolumns() <= col) {
			col = ss.getMaxcolumns() - 1;
		}
		ss.focusTo(row, col);
		getDesktopWorkbenchContext().getWorkbookCtrl().reGainFocus();
	}
	
	public void onOK() {
		focusOut = true;
		initVal = null;
		Utils.setEditText(ss.getSelectedSheet(), ss
				.getSelection().getTop(), ss.getSelection().getLeft(),
				getText());
		Position pos = ss.getCellFocus();
		int row = pos.getRow() + 1;
		int col = pos.getColumn();
		if (ss.getMaxrows() <= row) {
			row = ss.getMaxrows() - 1;
		}
		ss.focusTo(row, col);
		getDesktopWorkbenchContext().getWorkbookCtrl().reGainFocus();
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

	private DesktopWorkbenchContext getDesktopWorkbenchContext() {
		return DesktopWorkbenchContext.getInstance(Executions.getCurrent().getDesktop());
	}
}
