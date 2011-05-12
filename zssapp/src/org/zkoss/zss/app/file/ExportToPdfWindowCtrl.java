/* ExportToPdfWindowCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Sep 20, 2010 3:20:39 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.file;

import static org.zkoss.zss.app.base.Preconditions.checkNotNull;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.zkoss.poi.openxml4j.exceptions.InvalidFormatException;
import org.zkoss.poi.ss.usermodel.PrintSetup;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.poi.ss.util.AreaReference;
import org.zkoss.util.media.AMedia;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.zul.Dialog;
import org.zkoss.zss.app.zul.Zssapps;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Exporter;
import org.zkoss.zss.model.Exporters;
import org.zkoss.zss.model.impl.Headings;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Button;
import org.zkoss.zul.Checkbox;
import org.zkoss.zul.Filedownload;
import org.zkoss.zul.Radio;
import org.zkoss.zul.Radiogroup;
import org.zkoss.zul.Window;

/**
 * @author Sam
 *
 */
public class ExportToPdfWindowCtrl extends GenericForwardComposer {
	
	private Dialog _exportToPdfDialog;
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
	 * Original orientation setting
	 */
	boolean orgOrientation;
	Radio landscape;
	Radio portrait;
	
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
	
	Spreadsheet ss;
	
	public void onOpen$_exportToPdfDialog(ForwardEvent event) {
		loadPrintSetting();
		noHeader.setChecked(false);
		try {
			_exportToPdfDialog.setMode(Window.MODAL);
		} catch (InterruptedException e) {
		}
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
		loadOrientationSetting();
	}
	
	private void loadOrientationSetting() {
		//TODO: move to sheet context
		orgOrientation = ss.getSelectedSheet().getPrintSetup().getLandscape();
		if (ss.getSelectedSheet().getPrintSetup().getLandscape()) {
			orientation.setSelectedItem(landscape);
		}
		else
			orientation.setSelectedItem(portrait);
	}
	
	/**
	 * Apply the print setting to each page
	 */
	private void applyPrintSetting() {
		//TODO: move to sheet context
		ss.getSelectedSheet().getPrintSetup().setLandscape(orientation.getSelectedItem() == landscape);
		ss.getSelectedSheet().setPrintGridlines(includeGridlines());
		boolean isLandscape = orientation.getSelectedItem() == landscape;
		final Book book = ss.getBook(); 
		if (book == null) {
			return;
		}
		int numSheet = book.getNumberOfSheets();
		for (int i = 0; i < numSheet; i++) {
			Sheet sheet = ss.getSheet(i);
			PrintSetup setup = sheet.getPrintSetup();
			setup.setLandscape(isLandscape);
		}
	}
	
	private void revertPrintSetting() {
		final Book book = ss.getBook();
		if (book == null) {
			return;
		}
		int numSheet = book.getNumberOfSheets();
		for (int i = 0; i < numSheet; i++) {
			Sheet sheet = ss.getSheet(i);
			PrintSetup setup = sheet.getPrintSetup();
			setup.setLandscape(orgOrientation);
		}
	}

	public void onClick$export() 
		throws InvalidFormatException, IOException, InterruptedException {
		
		applyPrintSetting();
		
		Exporter c = Exporters.getExporter("pdf");
		if (c instanceof Headings) {
			((Headings)c).enableHeadings(includeHeadings());
		}
		
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		export(c, baos);
		
		final AMedia amedia = new AMedia("generatedReport.pdf", "pdf", "application/pdf", baos.toByteArray());

		Filedownload.save(amedia);
		
		revertPrintSetting();

		_exportToPdfDialog.fireOnClose(null);
	}
	
	private void export(Exporter exporter, OutputStream outputStream) {
		final Book book = ss.getBook();
		if (book == null) {
			return;
		}
		Radio seld = range.getSelectedItem();
		if (seld == allSheet) {
			exporter.export(book, outputStream);
		} else if (seld == currSelection){
			Rect rect = ss.getSelection();
			String area = ss.getColumntitle(rect.getLeft()) + ss.getRowtitle(rect.getTop()) + ":" + 
				ss.getColumntitle(rect.getRight()) + ss.getRowtitle(rect.getBottom());
			exporter.exportSelection(ss.getSelectedSheet(), new AreaReference(area), outputStream);
		} else
			exporter.export(ss.getSelectedSheet(), outputStream);
	}
	
	private boolean includeHeadings () {
		return !noHeader.isChecked();
	}
	
	private boolean includeGridlines(){
		return !noGridlines.isChecked();
	}
}