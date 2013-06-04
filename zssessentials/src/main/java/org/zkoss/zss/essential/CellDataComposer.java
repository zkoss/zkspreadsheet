package org.zkoss.zss.essential;

import java.io.IOException;
import java.util.ArrayList;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.SheetAnchor;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.CellData;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.essential.util.ClientUtil;
import org.zkoss.zss.ui.Position;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.DefaultUserAction;
import org.zkoss.zss.ui.event.AuxActionEvent;
import org.zkoss.zss.ui.event.CellAreaEvent;
import org.zkoss.zss.ui.event.CellEvent;
import org.zkoss.zss.ui.event.CellFilterEvent;
import org.zkoss.zss.ui.event.CellMouseEvent;
import org.zkoss.zss.ui.event.CellSelectionEvent;
import org.zkoss.zss.ui.event.CellSelectionUpdateEvent;
import org.zkoss.zss.ui.event.EditboxEditingEvent;
import org.zkoss.zss.ui.event.Events;
import org.zkoss.zss.ui.event.HeaderMouseEvent;
import org.zkoss.zss.ui.event.HeaderUpdateEvent;
import org.zkoss.zss.ui.event.CellHyperlinkEvent;
import org.zkoss.zss.ui.event.KeyEvent;
import org.zkoss.zss.ui.event.SheetDeleteEvent;
import org.zkoss.zss.ui.event.SheetEvent;
import org.zkoss.zss.ui.event.SheetSelectEvent;
import org.zkoss.zss.ui.event.StartEditingEvent;
import org.zkoss.zss.ui.event.StopEditingEvent;
import org.zkoss.zss.ui.event.WidgetKeyEvent;
import org.zkoss.zss.ui.event.WidgetUpdateEvent;
import org.zkoss.zul.Grid;
import org.zkoss.zul.Label;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Textbox;

/**
 * This class shows all the public ZK Spreadsheet you can listen to
 * @author dennis
 *
 */
public class CellDataComposer extends AbstractDemoComposer {

	@Wire
	Label cellType;
	@Wire
	Label cellFormatText;
	@Wire
	Label cellEditText;
	@Wire
	Label cellValue;
	@Wire
	Label cellResultType;
	@Wire
	Label cellRef;
	@Wire
	Textbox cellEditTextBox;
	
	
	@Listen("onCellFocused = #ss")
	public void onCellFocused(){
		Position pos = ss.getCellFocus();
		
		refreshCellInfo(pos.getRow(),pos.getColumn());		
	}
	
	private void refreshCellInfo(int row, int col){
		Range range = Ranges.range(ss.getSelectedSheet(),row,col);
		
		cellRef.setValue(Ranges.getCellReference(row, col));
		
		CellData data = range.getCellData();
		cellFormatText.setValue(data.getFormatText());
		cellEditText.setValue(data.getEditText());
		cellType.setValue(data.getType().toString());
		
		Object value = data.getValue();
		cellValue.setValue(value==null?"null":(value.getClass().getSimpleName()+" : "+value));
		cellResultType.setValue(data.getResultType().toString());
		
		cellEditTextBox.setValue(data.getEditText());
	}
	
	@Listen("onChange = #cellEditTextBox")
	public void onEditboxChange(){
		Position pos = ss.getCellFocus();
		Range range = Ranges.range(ss.getSelectedSheet(),pos.getRow(),pos.getColumn());
		CellData data = range.getCellData();
		if(data.validateEditText(cellEditTextBox.getValue())){
			data.setEditText(cellEditTextBox.getValue());
		}else{
			ClientUtil.showWarn("not a valid value");
		}
		refreshCellInfo(pos.getRow(),pos.getColumn());
		
	}
}



