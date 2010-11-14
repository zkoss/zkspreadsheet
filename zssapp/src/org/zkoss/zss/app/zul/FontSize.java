package org.zkoss.zss.app.zul;

import static org.zkoss.zss.app.base.Preconditions.checkNotNull;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.IdSpace;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zss.app.sheet.SheetHelper;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.CellEvent;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zul.Combobox;
import org.zkoss.zul.Div;

/**
 * 
 * @author sam
 *
 */
public class FontSize extends Div implements ZssappComponent, IdSpace {
	
	private final static String URI = "~./zssapp/html/fontSize.zul";
	
	//TODO: override combobox, provide onSelection event to change font Size
	private Combobox fontSizeCombobox;
	
	private Spreadsheet ss;

	public FontSize() {
		Executions.createComponents(URI, this, null);
		
		Components.wireVariables(this, this, '$', true, true);
		Components.addForwards(this, this, '$');
	}
	
	public void setText(String value) {
		fontSizeCombobox.setText(value);
	}
	
	public String getText() {
		return fontSizeCombobox.getText();
	}
	
	public void setWidth(String width) {
		fontSizeCombobox.setWidth(width);
	}
	
	
	public void onSelect$fontSizeCombobox(Event event) {
		Sheet sheet = ss.getSelectedSheet();
		short fontHeight = getFontHeight();
		Rect seldRect = SheetHelper.getSpreadsheetMaxSelection(ss);
		Utils.setFontHeight(sheet,
				seldRect, 
				fontHeight);

		setProperRowHeightByFontSize(sheet, seldRect);
		Events.postEvent(Events.ON_SELECT, this, event.getData());
		ZssappComponents.publishFontSizeChanged(ss, fontSizeCombobox.getSelectedItem().getLabel());
	}
	
	private void setProperRowHeightByFontSize(Sheet sheet, Rect rect) {
		int seldFontSize = Integer.parseInt(fontSizeCombobox.getText()); 
		
		int tRow = rect.getTop();
		int bRow = rect.getBottom();
		int col = rect.getLeft();
		
		for (int i = tRow; i <= bRow; i++) {
			//Note. the book helper return measured in twips (1/20 of  a point)
			if (seldFontSize > (BookHelper.getRowHeight(sheet, i) / 20)) {
				Ranges.range(sheet, i, col).setRowHeight(seldFontSize);
			}
		}
	}
	
	private short getFontHeight() {
		return (short) (Integer.parseInt(fontSizeCombobox.getText()) * 20);
	}
	
	@Override
	public Spreadsheet getSpreadsheet() {
		return ss;
	}

	@Override
	public void setSpreadsheet(Spreadsheet spreadsheet) {
		ss = checkNotNull(spreadsheet, "Spreadsheet is null");
		
		fontSizeCombobox.setWidgetListener("onOpen", "this.$f('" + ss.getId() + "', true).focus(false);");
		fontSizeCombobox.setWidgetListener("onSelect", "this.$f('" + ss.getId() + "', true).focus(false);");
		
		
		ZssappComponents.subscribeFontSizeChanged(ss, new EventListener() {
			
			@Override
			public void onEvent(Event event) throws Exception {
				if (fontSizeCombobox.isVisible()) {
					fontSizeCombobox.setText((String)event.getData());
				}
			}
		});
		
		ss.addEventListener(org.zkoss.zss.ui.event.Events.ON_CELL_FOUCSED, new EventListener(){

			@Override
			public void onEvent(Event event) throws Exception {
				CellEvent evt = (CellEvent)event;
				int row = evt.getRow();
				int col = evt.getColumn();
				Cell cell = Utils.getCell(ss.getSelectedSheet(), row, col);
				
				//set default
				fontSizeCombobox.setText("12");
				
				if (cell != null) {
					CellStyle cellStyle = cell.getCellStyle();
					int fontidx = cellStyle.getFontIndex();
					Font font = ss.getBook().getFontAt((short) fontidx);
					fontSizeCombobox.setText(Integer.toString(font.getFontHeightInPoints()));
				}
			}
		});
	}
}
