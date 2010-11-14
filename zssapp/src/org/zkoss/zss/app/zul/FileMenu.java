package org.zkoss.zss.app.zul;

import static org.zkoss.zss.app.base.Preconditions.checkNotNull;

import org.zkoss.util.resource.Labels;
import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.IdSpace;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zss.app.file.FileHelper;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Menu;
import org.zkoss.zul.Menuitem;
import org.zkoss.zul.Menupopup;

/**
 * 
 * @author sam
 *
 */
public class FileMenu extends Menu implements ZssappComponent, IdSpace {

	private final static String URI = "~./zssapp/html/fileMenu.zul";
	
	private Menupopup fileMenupopup;
	private Menuitem newFile;
	private Menuitem openFile;

	//TODO: not implement yet
	private Menuitem saveFile;
	private Menuitem saveFileAs;
	private Menuitem saveFileAndClose;
	//TODO: permission control
	private Menuitem deleteFile;
	private Menuitem importFile;
	private Menuitem exportToPdf;
	private Menuitem fileReversion;
	private Menuitem print;
	
	private Spreadsheet ss;
	
	public FileMenu() {
		Executions.createComponents(URI, this, null);
		
		Components.wireVariables(this, this, '$', true, true);
		Components.addForwards(this, this, '$');
		
		setLabel(Labels.getLabel("file"));
	}
	
	@Override
	public Spreadsheet getSpreadsheet() {
		return ss;
	}
	
	public void onClick$newFile() {
		FileHelper.openNewSpreadsheet(ss);
		
		
	}
	
	public void onClick$exportToPdf() {
		
	}

	@Override
	public void setSpreadsheet(Spreadsheet spreadsheet) {
		ss = checkNotNull(spreadsheet, "Spreadsheet is null");
		
		fileMenupopup.setWidgetListener(Events.ON_OPEN, "this.$f('" + ss.getId() + "', true).focus(false);");
	}

}
