/* ExportToPdfWindowCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Sep 20, 2010 3:20:39 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 3.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.app.file;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.UUID;

import org.zkoss.poi.openxml4j.exceptions.InvalidFormatException;
import org.zkoss.poi.ss.usermodel.PrintSetup;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.poi.ss.util.AreaReference;
import org.zkoss.io.Files;
import org.zkoss.util.media.AMedia;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.model.Exporter;
import org.zkoss.zss.model.impl.PdfExporter;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Button;
import org.zkoss.zul.Checkbox;
import org.zkoss.zul.Filedownload;
import org.zkoss.zul.Radio;
import org.zkoss.zul.Radiogroup;

/**
 * @author Sam
 *
 */
public class ExportToPdfWindowCtrl extends GenericForwardComposer {
	
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
	 * <p> Default: Include gridlines.
	 */
	Checkbox noGridlines;
	
	
	Button export;
	
	Spreadsheet ss;
	
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		ss = (Spreadsheet)getParam("spreadsheet");
		if (ss == null)
			throw new UiException("Spreadsheet object is null");

		loadPrintSetting();
	}
	
	private void loadPrintSetting() {
		range.setSelectedItem(currSheet);
		
		loadOrientationSetting();
	}
	
	private void loadOrientationSetting() {
		orgOrientation = ss.getSelectedSheet().getPrintSetup().getLandscape();
		if (ss.getSelectedSheet().getPrintSetup().getLandscape())
			orientation.setSelectedItem(landscape);
		else
			orientation.setSelectedItem(portrait);
	}
	
	/**
	 * Apply the print setting to each page
	 */
	private void applyPrintSetting() {
		ss.getSelectedSheet().getPrintSetup().setLandscape(orientation.getSelectedItem() == landscape);
		ss.getSelectedSheet().setPrintGridlines(includeGridlines());
		boolean isLandscape = orientation.getSelectedItem() == landscape;
		
		int numSheet = ss.getBook().getNumberOfSheets();
		for (int i = 0; i < numSheet; i++) {
			Sheet sheet = ss.getSheet(i);
			PrintSetup setup = sheet.getPrintSetup();
			setup.setLandscape(isLandscape);
		}
	}
	
	private void revertPrintSetting() {
		int numSheet = ss.getBook().getNumberOfSheets();
		for (int i = 0; i < numSheet; i++) {
			Sheet sheet = ss.getSheet(i);
			PrintSetup setup = sheet.getPrintSetup();
			setup.setLandscape(orgOrientation);
		}
	}
	
	private static Object getParam (String key) {
		return Executions.getCurrent().getArg().get(key);
	}

	public void onClick$export() 
		throws InvalidFormatException, IOException, InterruptedException {
		
		applyPrintSetting();
		
		Exporter c = new PdfExporter();
		((PdfExporter)c).enableHeadings(includeHeadings());
		
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		export(c, baos);
		
		final AMedia amedia = new AMedia("generatedReport.pdf", "pdf", "application/pdf", baos.toByteArray());

		Filedownload.save(amedia);
		
		revertPrintSetting();

		((Component)self.getSpaceOwner()).detach();
	}
	
	private void export(Exporter exporter, OutputStream outputStream) {
		Radio seld = range.getSelectedItem();
		if (seld == allSheet) {
			exporter.export(ss.getBook(), outputStream);	
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
