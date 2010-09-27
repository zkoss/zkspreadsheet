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
package org.zkoss.zss.app.export;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.UUID;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.zkoss.io.Files;
import org.zkoss.util.media.AMedia;
import org.zkoss.zhtml.Filedownload;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Exporter;
import org.zkoss.zss.model.impl.PdfExporter;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Button;
import org.zkoss.zul.Checkbox;
import org.zkoss.zul.Radiogroup;

/**
 * @author Sam
 *
 */
public class ExportToPdfWindowCtrl extends GenericForwardComposer {
	
	/**
	 * The range to export. All sheets, current sheet or selection range
	 * <p> Default: Export all sheets
	 */
	Radiogroup range;
	
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
	}
	
	private static Object getParam (String key) {
		return Executions.getCurrent().getArg().get(key);
	}

	public void onClick$export() 
		throws InvalidFormatException, IOException, InterruptedException{
		
		Exporter c = new PdfExporter();
		((PdfExporter)c).enableHeadings(includeHeadings());
		((PdfExporter)c).enableGridLines(includeGridlines());
		
		File tempDir = createTempDir();
		OutputStream os = null;
		
		Book wb = ss.getBook();
		String outputFilePath = tempDir.getAbsolutePath() 
	    	+ System.getProperty("file.separator") 
	    	+ System.currentTimeMillis() + ".pdf";
		
		os = new java.io.FileOutputStream(outputFilePath);
		c.export(wb, os);
		Files.close(os);
		
		final InputStream mediais = new FileInputStream(new File(outputFilePath));
		final AMedia amedia = new AMedia("generatedReport.pdf", "pdf", "application/pdf", mediais);

		Filedownload.save(amedia);
		
		//remove temp dir containing pdf output 
		recursiveDelete(tempDir);

		((Component)self.getSpaceOwner()).detach();
	}
	
	private boolean includeHeadings () {
		return !noHeader.isChecked();
	}
	
	private boolean includeGridlines(){
		return !noGridlines.isChecked();
	}
	
	/**
	 * Create a new temporary directory. Use something like
	 * {@link #recursiveDelete(File)} to clean this directory up since it isn't
	 * deleted automatically
	 * @return  the new directory
	 * @throws IOException if there is an error creating the temporary directory
	 */
	public File createTempDir() throws IOException {
		final File sysTempDir = new File(System.getProperty("java.io.tmpdir"));
		File newTempDir;
		final int maxAttempts = 9;
		int attemptCount = 0;
		do {
			attemptCount++;
			if (attemptCount > maxAttempts) {
				throw new IOException(
						"The highly improbable has occurred! Failed to "
								+ "create a unique temporary directory after "
								+ maxAttempts + " attempts.");
			}
			String dirName = UUID.randomUUID().toString();
			newTempDir = new File(sysTempDir, dirName);
		} while (newTempDir.exists());

		if (newTempDir.mkdirs()) {
			return newTempDir;
		} else {
			throw new IOException("Failed to create temp dir named "
					+ newTempDir.getAbsolutePath());
		}
	}
	
	/**
	 * Recursively delete file or directory
	 * @param fileOrDir
	 *          the file or dir to delete
	 * @return
	 *          true iff all files are successfully deleted
	 */
	public boolean recursiveDelete(File fileOrDir) {
		if (fileOrDir.isDirectory()) {
			// recursively delete contents
			for (File innerFile : fileOrDir.listFiles()) {
				if (!recursiveDelete(innerFile)) {
					return false;
				}
			}
		}
		return fileOrDir.delete();
	}
}
