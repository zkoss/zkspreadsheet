/* FileMenu.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 23, 2010 5:43:38 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul;

import java.io.ByteArrayOutputStream;

import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.IdSpace;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.app.file.FileHelper;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zss.app.zul.ctrl.WorkbookCtrl;
import org.zkoss.zss.app.zul.ctrl.WorkspaceContext;
import org.zkoss.zul.Filedownload;
import org.zkoss.zul.Menu;
import org.zkoss.zul.Menuitem;
import org.zkoss.zul.Menupopup;

/**
 * 
 * @author sam
 *
 */
public class FileMenu extends Menu implements IdSpace {
	
	private Menupopup fileMenupopup;
	private Menuitem newFile;
	private Menuitem openFile;

	
	private Menuitem saveFile;
	//TODO: not implement yet
	private Menuitem saveFileAs;
	private Menuitem saveFileAndClose;
	//TODO: permission control
	private Menuitem deleteFile;
	private Menuitem importFile;
	private Menuitem exportToPdf;
	private Menuitem exportToExcel;
	private Menuitem fileReversion;
	private Menuitem print;

	
	public FileMenu() {
		Executions.createComponents(Consts._FileMenu_zul, this, null);
		Components.wireVariables(this, this, '$', true, true);
		Components.addForwards(this, this, '$');

		importFile.setDisabled(!FileHelper.hasImportPermission());
		
		boolean saveDisabled = !FileHelper.hasSavePermission();
		saveFile.setDisabled(saveDisabled);
		//TODO: save as not implement yet
		saveFileAs.setDisabled(true);
		saveFileAndClose.setDisabled(saveDisabled);
	}
	
	public void setExportDisabled(boolean disabled) {
		
		if (FileHelper.hasSavePermission()) {
			saveFile.setDisabled(disabled);
			//TODO: not implemented yet
			saveFileAs.setDisabled(true);
			saveFileAndClose.setDisabled(disabled);
		}
		
		deleteFile.setDisabled(disabled);
		exportToPdf.setDisabled(disabled);
		exportToExcel.setDisabled(disabled);
		
		//TODO: not implemented yet
		fileReversion.setDisabled(true);
		print.setDisabled(true);
	}
	
	public void onClick$newFile() {
		getDesktopWorkbenchContext().getWorkbookCtrl().newBook();
		getDesktopWorkbenchContext().fireWorkbookChanged();
	}
	
	public void onClick$openFile() {
		getDesktopWorkbenchContext().getWorkbenchCtrl().openOpenFileDialog();
	}
	
	public void onClick$saveFile() {
		//TODO: refactor duplicate save logic
		DesktopWorkbenchContext workbench = getDesktopWorkbenchContext();
		if (workbench.getWorkbookCtrl().hasFileName()) {
			workbench.getWorkbookCtrl().save();
			workbench.fireWorkbookSaved();
		} else
			workbench.getWorkbenchCtrl().openSaveFileDialog();
	}
	
	public void onClick$saveFileAs() {
		throw new UiException("save is not implement yet");
	}
	
	public void onClick$saveFileAndClose() {
		//TODO: refactor duplicate save logic
		DesktopWorkbenchContext workbench = getDesktopWorkbenchContext();
		if (workbench.getWorkbookCtrl().hasFileName()) {
			workbench.getWorkbookCtrl().save();
			workbench.getWorkbookCtrl().close();
			workbench.fireWorkbookSaved();
			workbench.fireWorkbookChanged();
		} else
			workbench.getWorkbenchCtrl().openSaveFileDialog();
	}
	
	public void onClick$deleteFile() {
		DesktopWorkbenchContext workbench = getDesktopWorkbenchContext();
		if(!workbench.getWorkbookCtrl().hasFileName()) {
			workbench.getWorkbookCtrl().close();
			workbench.fireWorkbookChanged();
			return;
		}
		
		WorkspaceContext.getInstance(Executions.getCurrent().getDesktop()).
			delete(workbench.getWorkbookCtrl().getSrc());
		workbench.getWorkbookCtrl().close();
		workbench.fireWorkbookChanged();
	}
	
	public void onClick$importFile() {
		getDesktopWorkbenchContext().getWorkbenchCtrl().openImportFileDialog();
	}

	public void setExportPdfDisabled(boolean disabled) {
		exportToPdf.setDisabled(disabled);
	}
	
	public void onClick$exportToPdf() {
		getDesktopWorkbenchContext().getWorkbenchCtrl().openExportPdfDialog();
	}

	public void setExportExcelDisabled(boolean disabled) {
		exportToExcel.setDisabled(disabled);
	}
	
	public void onClick$exportToExcel() {
		WorkbookCtrl bookCtrl = getDesktopWorkbenchContext().getWorkbookCtrl();
		ByteArrayOutputStream out = bookCtrl.exportToExcel();
		Filedownload.save(out.toByteArray(), "application/file", bookCtrl.getBookName());
	}
	
	public void onClick$fileReversion() {
		throw new UiException("reversion not implement yet");
	}
	
	public void onClick$print() {
		throw new UiException("print not implement yet");
	}
	
	public void onOpen$fileMenupopup() {
		getDesktopWorkbenchContext().getWorkbookCtrl().reGainFocus();
	}
	
	protected DesktopWorkbenchContext getDesktopWorkbenchContext() {
		return Zssapp.getDesktopWorkbenchContext(this);
	}
}