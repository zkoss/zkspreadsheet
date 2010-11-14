/* CellMenu.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 11, 2010 3:06:55 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.app.zul;

import static org.zkoss.zss.app.base.Preconditions.checkNotNull;

import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.IdSpace;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.event.MouseEvent;
import org.zkoss.zss.app.cell.EditHelper;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Menuitem;
import org.zkoss.zul.Menupopup;
/**
 * @author Sam
 *
 */
public class CellMenupopup extends Menupopup implements ZssappComponent, IdSpace {
	
	private final static String URI = "~./zssapp/html/cellMenu.zul";
	
	private final static String KEY_INSERT_WINDOW_DIALOG = "org.zkoss.zss.app.zul.cellMenupopup.insertWindowDialog";
	private final static String KEY_DELETE_WINDOW_DIALOG = "org.zkoss.zss.app.zul.cellMenupopup.deleteWindowDialog";
	
	/* Setting */
	//private boolean _insertVisible; 
	//private boolean _insertDisable;
	
	/* Components */
	private Menuitem cut;
	private Menuitem copy;
	private Menuitem paste;
	private Menuitem pasteSpecial;
	
	private Menuitem insert;
	private Menuitem delete;
	private Menuitem clearContent;
	private Menuitem clearStyle;
	
	//TODO: not implement yet
	//private Menuitem filter
	private Menuitem sort;
	
	//TODO: not implement yet
	//private Menuitem comment
	
	//Others
	private Menuitem formula;
	//TODO: not implement yet
	//private Menuitem chart;
	//private Menuitem pic;
	private Menuitem format;
	private Menuitem hyperlink;
	//TODO: not implement yet
	//private Menuitem removeHyperlink;
	
	
	private Spreadsheet ss;
	
	public CellMenupopup() {
		Executions.createComponents(URI, this, null);
		Components.wireVariables(this, this, '$', true, true);
		Components.addForwards(this, this, '$');
	}
	
	//TODO: CellMenupopup API
//	public void setInsertVisible(boolean visible) {
//	}
//	
//	public void isInsertVisible() {
//	}
//	
//	public void setInsertDisable(boolean disable) {
//		_insertDisable = disable;
//		
//	}
//	
//	public boolean isInsertDisable() {
//		return _insertDisable;
//	}
	
	
	public void setSpreadsheet(Spreadsheet spreadsheet) {
		ss = checkNotNull(spreadsheet, "Spreadsheet is null");
		setWidgetListener("onOpen", "this.$f('" + ss.getId() + "', true).focus(false);");
	}
	
	public Spreadsheet getSpreadsheet() {
		return ss;
	}
	
	public void onClick$cut() {
		EditHelper.doCut(ss);
	}
	
	public void onClick$copy() {
		EditHelper.doCopy(ss);
	}
	
	public void onClick$paste() {
		EditHelper.doPaste(ss);
	}
	
	public void onClick$pasteSpecial(MouseEvent event) {
		Executions.createComponents(
				"~./zssapp/html/pasteSpecialWindowDlg.zul", null, ZssappComponents.newSpreadsheetArg(ss));
	}
	
	public void onClick$insert(ForwardEvent event) {
		MouseEvent evt = (MouseEvent)event.getOrigin();
		
		InsertWindowDialog dialog = (InsertWindowDialog)ss.getAttribute(KEY_INSERT_WINDOW_DIALOG);
		if (dialog == null) {
			dialog = new InsertWindowDialog();
			dialog.setSpreadsheet(ss);
		}
		//TODO: calculate proper dialog position
		dialog.setLeft(evt.getX() + "px");
		dialog.setTop(evt.getY() + "px");
		dialog.doPopup();
	}
	
	public void onClick$delete(ForwardEvent event) {
		MouseEvent evt = (MouseEvent)event.getOrigin();
		
		DeleteWindowDialog dialog = (DeleteWindowDialog)ss.getAttribute(KEY_DELETE_WINDOW_DIALOG);
		if (dialog == null) {
			dialog = new DeleteWindowDialog();
			dialog.setSpreadsheet(ss);
		}
		//TODO: calculate proper dialog position
		dialog.setLeft(evt.getX() + "px");
		dialog.setTop(evt.getY() + "px");
		dialog.doPopup();
	}
}