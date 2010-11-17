/* FontFamily.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 7, 2010 10:31:11 AM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.app.zul;

import static org.zkoss.zss.app.base.Preconditions.checkNotNull;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.IdSpace;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zss.app.sheet.SheetHelper;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.CellEvent;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zul.Combobox;
import org.zkoss.zul.Div;

/**
 * @author Sam
 *
 */
public class FontFamily extends Div implements ZssappComponent, IdSpace{
	
	private final static String URI = "~./zssapp/html/fontFamily.zul";
	
	//TODO: override combobox, provide onSelection event, 
	// change font size when onSelection, this shall be configanle
	private Combobox fontfamilyCombobox;
	
	private Spreadsheet ss;
	
	
	public FontFamily() {
		Executions.createComponents(URI, this, null);
		
		Components.wireVariables(this, this, '$', true, true);
		Components.addForwards(this, this, '$');
	}
	
	public void onSelect$fontfamilyCombobox(Event event) {
		String font = fontfamilyCombobox.getSelectedItem().getLabel();
		Utils.setFontFamily(ss.getSelectedSheet(), SheetHelper.getSpreadsheetMaxSelection(ss), font);
		//post event out for container
		Events.postEvent(Events.ON_SELECT, this, event.getData());
		//publish event queue for desktop
		ZssappComponents.publishFontFamilyChanged(ss, font);
	}
	
	public String getText() {
		return fontfamilyCombobox.getText();
	}
	
	public void setText(String fontFamily) {
		fontfamilyCombobox.setText(fontFamily);
	}

	public void setWidth(String width) {
		fontfamilyCombobox.setWidth(width);
	}
	
	public void setStyle(String style) {
		fontfamilyCombobox.setStyle(style);
	}

	@Override
	public Spreadsheet getSpreadsheet() {
		return ss;
	}

	@Override
	public void setSpreadsheet(Spreadsheet spreadsheet) {
		ss = checkNotNull(spreadsheet, "Spreadsheet is null");
		initFontFamily();
	}

	private void initFontFamily() {
		//TODO: add event queue
		fontfamilyCombobox.setWidgetListener(Events.ON_OPEN, "this.$f('" + ss.getId() + "', true).focus(false);");
		fontfamilyCombobox.setWidgetListener(Events.ON_SELECT, "this.$f('" + ss.getId() + "', true).focus(false);");
		
		ZssappComponents.subscribeFontFamilyChanged(ss, new EventListener() {
			
			@Override
			public void onEvent(Event event) throws Exception {
				if (fontfamilyCombobox.isVisible())
					fontfamilyCombobox.setText((String)event.getData());
			}
		});
		
		ss.addEventListener(org.zkoss.zss.ui.event.Events.ON_CELL_FOUCSED, new EventListener() {
			public void onEvent(Event event) throws Exception {
				final Book book = ss.getBook();
				if (book == null) {
					return;
				}
				CellEvent evt = (CellEvent)event;
				int row = evt.getRow();
				int col = evt.getColumn();
				Cell cell = Utils.getCell(ss.getSelectedSheet(), row, col);
				
				//set to default font family
				fontfamilyCombobox.setText("Calibri");
				
				if (cell != null) {
					CellStyle cellStyle = cell.getCellStyle();
					Font font = book.getFontAt(cellStyle.getFontIndex());
					fontfamilyCombobox.setText(font.getFontName());
				}
			}
		});
	}
}