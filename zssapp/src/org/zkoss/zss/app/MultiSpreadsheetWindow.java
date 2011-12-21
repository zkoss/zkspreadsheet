package org.zkoss.zss.app;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.ext.AfterCompose;
//import org.zkoss.zss.model.Cell;
//import org.zkoss.zss.model.Sheet;
//import org.zkoss.zss.model.impl.SheetImpl;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.CellEvent;
import org.zkoss.zss.ui.event.Events;
import org.zkoss.zss.ui.event.StopEditingEvent;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zul.Window;

public class MultiSpreadsheetWindow extends Window implements AfterCompose{
	Spreadsheet[] ss;
	
	public void afterCompose() {
		ss=new Spreadsheet[2];
		ss[0]=(Spreadsheet) getFellow("ss0");
		ss[1]=(Spreadsheet) getFellow("ss1");
		for(int i=0; i<2;i++){
			if(ss[i]==null)
				continue;
			
			ss[i].addEventListener(Events.ON_STOP_EDITING, new EventListener() {
				public void onEvent(Event event) throws Exception {
					onStopEditingEvent((StopEditingEvent) event);
				}
			});
			ss[i].addEventListener(Events.ON_CELL_FOUCSED,
					new EventListener() {
				public void onEvent(Event event) throws Exception {
					onFocusedEvent((CellEvent) event);
				}
			});
			
		}
	}
	
	public void onStopEditingEvent(StopEditingEvent event){
		int row=event.getRow();
		int col=event.getColumn();
		
		Worksheet targetSheet;
		if(event.getTarget()==ss[0])
			targetSheet=ss[1].getSelectedSheet();
		else
			targetSheet=ss[0].getSelectedSheet();
		
		Cell tmpCell=Utils.getOrCreateCell(targetSheet, row, col);
		Utils.setEditText(tmpCell, (String)event.getEditingValue());
	}
	
	public void onFocusedEvent(CellEvent event){
		int row=event.getRow();
		int col=event.getColumn();
		
		Spreadsheet targetSpreadsheet;
		if(event.getTarget()==ss[0])
			ss[1].moveEditorFocus("ss1", "ss1", "red", row, col);
		else
			ss[0].moveEditorFocus("ss0", "ss0", "green", row, col);
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
}