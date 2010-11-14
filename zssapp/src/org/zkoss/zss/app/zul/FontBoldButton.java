package org.zkoss.zss.app.zul;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.util.resource.Labels;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zss.app.sheet.SheetHelper;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.CellEvent;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zul.Toolbarbutton;

public class FontBoldButton extends Toolbarbutton implements ZssappComponent {

	
	private Spreadsheet ss;
	
	private boolean _isBold;
	
	public FontBoldButton() {
		setImage("~./zssapp/image/edit-bold.png");
		setTooltip(Labels.getLabel("font.bold'"));
	}
	
	public void onClick() {
		Utils.setFontBold(ss.getSelectedSheet(), 
				SheetHelper.getSpreadsheetMaxSelection(ss), 
				_isBold = !_isBold);
		setSclass(_isBold ? "clicked" : "");
		ZssappComponents.publishFontBoldChanged(ss, _isBold);
	}

	@Override
	public Spreadsheet getSpreadsheet() {
		return ss;
	}
	
	public void setBold(boolean isBold) {
		Utils.setFontBold(ss.getSelectedSheet(), 
				SheetHelper.getSpreadsheetMaxSelection(ss), 
				isBold);
		_isBold = isBold;
		setSclass(_isBold ? "clicked" : "");
	}

	@Override
	public void setSpreadsheet(Spreadsheet spreadsheet) {
		ss = spreadsheet;
		setWidgetListener("onClick", "this.$f('" + ss.getId() + "', true).focus(false);");
		
		ZssappComponents.subscribeFontBoldChanged(ss, new EventListener() {
			
			@Override
			public void onEvent(Event event) throws Exception {
				if (isVisible()) {
					_isBold = (Boolean)event.getData();
					setSclass(_isBold ? "clicked" : "");
				}
			}
		});
		
		ss.addEventListener(org.zkoss.zss.ui.event.Events.ON_CELL_FOUCSED, new EventListener() {
			
			@Override
			public void onEvent(Event event) throws Exception {
				CellEvent evt = (CellEvent)event;
				int row = evt.getRow();
				int col = evt.getColumn();
				Cell cell = Utils.getCell(ss.getSelectedSheet(), row, col);
				
				//set default
				setSclass("");
				
				if (cell != null) {
					CellStyle cellStyle = cell.getCellStyle();
					Font font = ss.getBook().getFontAt(cellStyle.getFontIndex());
					_isBold = font.getBoldweight() == Font.BOLDWEIGHT_BOLD;
					if (_isBold)
						setSclass("clicked");
				}
			}
		});
	}
}