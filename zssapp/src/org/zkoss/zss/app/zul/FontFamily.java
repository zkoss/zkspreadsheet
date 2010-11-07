package org.zkoss.zss.app.zul;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zss.app.MainWindowCtrl;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.CellEvent;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zul.Combobox;
import org.zkoss.zul.Div;

public class FontFamily extends Div {
	
	private Combobox fontfamilyCombobox;
	
	private Spreadsheet ss;
	
	public void setUri(String uri) {
		Executions.createComponents(uri, this, null);
		Components.wireVariables(this, (Object)this);
		Components.addForwards(this, (Object)this, '$');
	}
	
	public void onSelect$fontfamilyCombobox(Event event) {

		String font = fontfamilyCombobox.getSelectedItem().getLabel();
		Events.postEvent(Events.ON_SELECT, this, event.getData());
		MainWindowCtrl.getInstance().setFontFamily(font);
	}
	
	public String getText() {
		return fontfamilyCombobox.getText();
	}
	
	public void setText(String fontFamily) {
		fontfamilyCombobox.setText(fontFamily);
	}
	
	public void onCreate() {
		ss = MainWindowCtrl.getInstance().getSpreadsheet();
		
		ss.addEventListener(org.zkoss.zss.ui.event.Events.ON_CELL_FOUCSED, new EventListener() {
			public void onEvent(Event event) throws Exception {
				CellEvent evt = (CellEvent)event;
				int row = evt.getRow();
				int col = evt.getColumn();

				Cell cell = Utils.getCell(ss.getSelectedSheet(), row, col);
				
				//set to default font family
				fontfamilyCombobox.setText("Calibri");
				
				if (cell != null) {
					CellStyle cellStyle = cell.getCellStyle();
					Font font = ss.getBook().getFontAt(cellStyle.getFontIndex());
					fontfamilyCombobox.setText(font.getFontName());
				}
			}
		});
	}

	public void setWidth(String width) {
		fontfamilyCombobox.setWidth(width);
	}
}