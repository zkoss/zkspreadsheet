package org.zkoss.zss.app.zul;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zss.app.MainWindowCtrl;
import org.zkoss.zss.app.cell.CellHelper;
import org.zkoss.zss.app.zul.api.Colorbutton;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.CellMouseEvent;
import org.zkoss.zss.ui.event.Events;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zul.Combobox;
import org.zkoss.zul.Toolbarbutton;
import org.zkoss.zul.Window;

public class CellContext extends Window {
	
	FontFamily fontfamily;
	Combobox _fontSizeCombobox;
	Toolbarbutton _boldBtn;
	Toolbarbutton _italicBtn;
	Toolbarbutton _alignLeftBtn;
	Toolbarbutton _alignCenterBtn;
	Toolbarbutton _alignRightBtn;
	Colorbutton _fontColorBtn;
	Colorbutton _backgroundColorBtn;
	Toolbarbutton _borderBtn;
	Toolbarbutton _mergeCellBtn;

	
	private Spreadsheet ss;
	
	public CellContext() {
		setVisible(false);
		setSclass("fastIconWin");
		setVflex("min");
		setWidth("210px");
	}
	
	public void setUri(String uri) {
		Executions.createComponents(uri, this, null);
		Components.wireVariables(this, (Object)this);
		Components.addForwards(this, (Object)this, '$');
	}
	
	public void onCreate() {
		ss = MainWindowCtrl.getInstance().getSpreadsheet();

		_fontColorBtn = (Colorbutton)this.getFellow("_fontColorBtn");
		_backgroundColorBtn = (Colorbutton)this.getFellow("_backgroundColorBtn");
		
		
		ss.addEventListener(Events.ON_CELL_RIGHT_CLICK, 
			new EventListener() {
				public void onEvent(Event event) throws Exception {
					CellMouseEvent evt = (CellMouseEvent)event;
					int clientX = evt.getClientx();
					int clientY = evt.getClienty();

					setIconsAttributes(
						Utils.getCell(
							ss.getSelectedSheet(), 
							ss.getSelection().getTop(), 
							ss.getSelection().getLeft()));
					
					CellContext.this.setLeft(Integer.toString(clientX + 5) + "px");
					CellContext.this.setTop(Integer.toString(clientY - 100) + "px");
					CellContext.this.doPopup();					
				}
		});
	}
	
	private void setIconsAttributes(Cell cell) {
		//Sets default
		fontfamily.setText("Calibri");
		_fontColorBtn.setColor("#000000");
		_backgroundColorBtn.setColor("#FFFFFF");
		if (cell == null)
			return;
		
		Font font = CellHelper.getFont(cell);
		fontfamily.setText(font.getFontName());
		_fontSizeCombobox.setText(Integer.toString(font.getFontHeightInPoints()));
		_boldBtn.setSclass(CellHelper.isBold(font) ? "clicked" : "");
		_boldBtn.setClass(font.getItalic() ? "clicked" : "");
		_alignLeftBtn.setSclass(CellHelper.isAlignLeft(cell) ? "clicked" : "");
		_alignCenterBtn.setSclass(CellHelper.isAlignCenter(cell) ? "clicked" : "");
		_alignRightBtn.setSclass(CellHelper.isAlignRight(cell) ? "clicked" : "");
		_fontColorBtn.setColor(CellHelper.getFontHTMLColor(cell, font));
		_backgroundColorBtn.setColor(CellHelper.getBackgroundHTMLColor(cell));
	}
	
	public void onSelect$fontfamily() {
		this.setVisible(false);
	}
	
	public void onChange$_fontColorBtn(Event event) {
		MainWindowCtrl.getInstance().setFontColor(_fontColorBtn.getColor());
		setVisible(false);
	}

	public void onChange$_backgroundColorBtn(Event event) {
		MainWindowCtrl.getInstance().setBackgroundColor(_backgroundColorBtn.getColor());
		setVisible(false);
	}

	public void onSelect$_fontSizeCombobox(Event event) {
		setVisible(false);
		MainWindowCtrl.getInstance().setFontSize(_fontSizeCombobox.getSelectedItem().getLabel());
	}

	public void onBorderSelector(ForwardEvent evt) {
		MainWindowCtrl.getInstance().onBorderSelector(evt);
		setVisible(false);
	}
}
