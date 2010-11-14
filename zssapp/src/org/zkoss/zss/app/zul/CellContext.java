/* CellContext.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 7, 2010 10:31:11 AM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 3.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.app.zul;

import static org.zkoss.zss.app.base.Preconditions.checkNotNull;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.IdSpace;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zss.app.cell.CellHelper;
import org.zkoss.zss.app.zul.api.Colorbutton;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.CellMouseEvent;
import org.zkoss.zss.ui.event.Events;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zul.Toolbarbutton;
import org.zkoss.zul.Window;

/**
 * @author Sam
 *
 */
public class CellContext extends Window implements ZssappComponent, IdSpace {
	
	private final static String URI = "~./zssapp/html/cellContext.zul";
	
	FontFamily fontfamily;
	FontSize fontSize;
	//Toolbarbutton _boldBtn;
	FontBoldButton boldBtn;
	Toolbarbutton _italicBtn;
	Toolbarbutton _alignLeftBtn;
	Toolbarbutton _alignCenterBtn;
	Toolbarbutton _alignRightBtn;
	Colorbutton _fontColorBtn;
	Colorbutton _backgroundColorBtn;
	Borderbutton borderBtn;
	Toolbarbutton _mergeCellBtn;
	
	private Spreadsheet ss;
	
	public CellContext() {
		Executions.createComponents(URI, this, null);
		
		Components.wireVariables(this, this, '$', true, true);
		Components.addForwards(this, this, '$');
		setVisible(false);
		setSclass("fastIconWin");
		setVflex("min");
		setWidth("210px");
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
		fontSize.setText(Integer.toString(font.getFontHeightInPoints()));
		//_boldBtn.setSclass(CellHelper.isBold(font) ? "clicked" : "");
		//_boldBtn.setClass(font.getItalic() ? "clicked" : "");
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
		//TODO color button shall do event, should not call controller
		Zssapp.setFontColor(ss, _fontColorBtn.getColor());
		setVisible(false);
	}

	public void onChange$_backgroundColorBtn(Event event) {
		//TODO color button shall do event, should not call controller
		Zssapp.setBackgroundColor(ss, _backgroundColorBtn.getColor());
		setVisible(false);
	}

	public void onSelect$fontSize() {
		setVisible(false);
	}

	public void onClick$borderBtn() {
		setVisible(false);
	}
	
	public void onClick$boldBtn() {
		setVisible(false);
	}

	@Override
	public Spreadsheet getSpreadsheet() {
		return ss;
	}

	@Override
	public void setSpreadsheet(Spreadsheet spreadsheet) {
		ss = checkNotNull(spreadsheet, "Spreadsheet is null");
		initCellContext();
	}
	
	private void initCellContext() {
		//move this to util
		fontSize.setSpreadsheet(ss);
		fontfamily.setSpreadsheet(ss);
		borderBtn.setSpreadsheet(ss);
		boldBtn.setSpreadsheet(ss);
		
		setWidgetListener("onShow", "this.$f('" + ss.getId() + "', true).focus(false);");

		/*Note. the colorbutton here is interface, need to use getFellow to bind field*/
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
}
