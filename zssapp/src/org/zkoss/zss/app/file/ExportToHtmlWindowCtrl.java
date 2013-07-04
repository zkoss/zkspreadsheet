/* ExportToHtmlWindowCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		March 05, 2011 3:20:39 PM , Created by Peter
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.file;

import static org.zkoss.zss.app.base.Preconditions.checkNotNull;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.zkoss.poi.ss.usermodel.PrintSetup;
import org.zkoss.util.media.AMedia;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.app.zul.Dialog;
import org.zkoss.zss.app.zul.Zssapps;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Button;
import org.zkoss.zul.Checkbox;
import org.zkoss.zul.Filedownload;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.Radio;
import org.zkoss.zul.Radiogroup;
import org.zkoss.zul.Window;

/**
 * @author Peter
 *
 */
public class ExportToHtmlWindowCtrl extends GenericForwardComposer {
	
	private Dialog _exportToHtmlDialog;
	/**
	 * The range to export. All sheets, current sheet or selection range
	 * <p> Default: Selected sheet
	 */
	Radiogroup range;
	Radio currSelection;
	Radio currSheet;
	Radio allSheet;
	
	/**
	 * The document's orientation. Landscape or Portrait
	 */
	Radiogroup orientation;
	
	/**
	 * Indicate whether include header or not
	 * <p> Default: Include header
	 */
	Checkbox noHeader;
	
	/**
	 * Indicate whether include gridlines or not.
	 */
	Checkbox noGridlines;
	
	Button export;
		
	Rect selection;
	Spreadsheet ss;
	
	public void onOpen$_exportToHtmlDialog(ForwardEvent event) {
		loadPrintSetting();
		
		selection = (Rect) event.getOrigin().getData();
		currSelection.setDisabled(selection == null);
		
		noHeader.setChecked(false);
		_exportToHtmlDialog.setMode(Window.MODAL);
	}
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		//TODO: use event, don't send arg
		ss = checkNotNull(Zssapps.getSpreadsheetFromArg(), "Spreadsheet is null");
	}
	
	private void loadPrintSetting() {
		noGridlines.setChecked(!ss.getSelectedSheet().isPrintGridlines());
		range.setSelectedItem(currSheet);
	}
	
	
	/**
	 * Apply the print setting to each page
	 */
	private void applyPrintSetting() {
		//TODO: move to sheet context
		ss.getSelectedSheet().getPoiSheet().setPrintGridlines(includeGridlines());//TODO shouldn't this go through Range API?
		final Book book = ss.getBook(); 
		if (book == null) {
			return;
		}
		int numSheet = book.getNumberOfSheets();
		for (int i = 0; i < numSheet; i++) {
			Sheet sheet = book.getSheetAt(i);
			org.zkoss.poi.ss.usermodel.Sheet poiSheet = sheet.getPoiSheet();
			PrintSetup setup = poiSheet.getPrintSetup();
		}
	}
	
	private void revertPrintSetting() {
		final Book book = ss.getBook();
		if (book == null) {
			return;
		}
		int numSheet = book.getNumberOfSheets();
		for (int i = 0; i < numSheet; i++) {
			Sheet sheet = book.getSheetAt(i);
			org.zkoss.poi.ss.usermodel.Sheet poiSheet = sheet.getPoiSheet();
			PrintSetup setup = poiSheet.getPrintSetup();
		}
	}

	public void onClick$export() 
		throws IOException, InterruptedException {
		
		applyPrintSetting();
		
		Exporter c = Exporters.getExporter("html");
		if(c==null){
			Messagebox.show("can't find html exporter");
			return;
		}
//		if (c.isSupportHeadings()) {
//			c.enableHeadings(includeHeadings());
//		}
		
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		export(c, baos);
		
		//test
//		System.out.println(baos.toString());
		
		final AMedia amedia = new AMedia("generatedReport.html", "html", "text/html", baos.toByteArray());

		Filedownload.save(amedia);
		
		revertPrintSetting();

		_exportToHtmlDialog.fireOnClose(null);
	}
	
	//test only
//	public void onClick$test(){
//		Sheet sheet = ss.getBook().getSheetAt(0);
//		String expr="=A1";
//		Ranges.range((Worksheet) sheet,3,3).setValue(expr);
//	}
	
	private void export(Exporter exporter, OutputStream outputStream) throws IOException{
		final Book book = ss.getBook();
		if (book == null) {
			return;
		}
		Radio seld = range.getSelectedItem();
		if (seld == allSheet) {
			exporter.export(book, outputStream);
		} else if (seld == currSelection){
			Rect selection = ss.getSelection();
			exporter.export(ss.getSelectedSheet(), selection, outputStream);
		} else {
			exporter.export(ss.getSelectedSheet(), outputStream);
		}
	}
	
	private boolean includeHeadings () {
		return !noHeader.isChecked();
	}
	
	private boolean includeGridlines(){
		return !noGridlines.isChecked();
	}
}