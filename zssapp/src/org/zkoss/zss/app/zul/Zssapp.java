/* Zssapp.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 16, 2010 6:22:52 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul;

import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.IdSpace;
import org.zkoss.zss.app.MainWindowCtrl;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Div;
import org.zkoss.zul.Window;

/**
 * 
 * @author sam
 *
 */
public class Zssapp extends Div implements IdSpace  {
	
	private final static String URI = "~./zssapp/html/zssapp.zul";
	
	private final static String KEY_ZSSAPP = "org.zkoss.zss.app.zul.zssApp";
	/*Default spreadsheet*/
	//TODO: in mainWin id space, move to here
	private Spreadsheet spreadsheet;
	
	//TODO: remove mainWin and mainWinCtrl, use Zssapp as controller
	private Window mainWin;
	MainWindowCtrl mainWinCtrl;
	
	
	public Zssapp() {
		Executions.createComponents(URI, this, null);
		Components.wireVariables(this, this, '$', true, true);
		Components.addForwards(this, this, '$');
		
		spreadsheet = (Spreadsheet)mainWin.getFellow("spreadsheet");
		mainWinCtrl = (MainWindowCtrl)mainWin.getAttribute(mainWin.getId() + "$composer");
		bindAppController(spreadsheet, mainWinCtrl);
		spreadsheet.setAttribute(KEY_ZSSAPP, this);
		
	}
	
	public void setSrc(String src) {
		spreadsheet.setSrc(src);
		mainWinCtrl.redrawSheetTabbox();
	}
	
	public void setMaxrows(int maxrows) {
		spreadsheet.setMaxrows(maxrows);
	}
	
	public void setMaxcolumns(int maxcols) {
		spreadsheet.setMaxcolumns(maxcols);
	}

	public void setWidth(String width) {
		super.setWidth(width);
		mainWin.setWidth(width);
	}
	
	public void setHeight(String height) {
		super.setHeight(height);
		mainWin.setHeight(height);
	}
	
	public void setHflex(String flex) {
		super.setHflex(flex);
		mainWin.setHflex(flex);
	}
	
	public void setVflex(String flex) {
		super.setVflex(flex);
		mainWin.setVflex(flex);
	}
	
	public Spreadsheet getSpreadsheet() {
		return spreadsheet;
	} 
	
	/**
	 * TODO: replace MainWindowCtrl, use this object as controller
	 */
	
	private static String KEY_ZSSAPP_CONTROLLER = "org.zkoss.zss.app.zul.zssapp.appController";
	private void bindAppController(Spreadsheet ss, Object ctrl) {
		ss.setAttribute(KEY_ZSSAPP_CONTROLLER, ctrl);
	}
	
	private static Object getAppController(Spreadsheet ss) {
		return ss.getAttribute(KEY_ZSSAPP_CONTROLLER);
	}
	
	public static Zssapp getInstance(Spreadsheet spreadsheet) {
		return (Zssapp)spreadsheet.getAttribute(KEY_ZSSAPP);
	}
}