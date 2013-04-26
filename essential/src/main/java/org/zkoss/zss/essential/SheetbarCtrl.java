package org.zkoss.zss.essential;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.Map;

import org.dom4j.IllegalAddException;
import org.zkoss.image.AImage;
import org.zkoss.poi.ss.usermodel.Workbook;
import org.zkoss.util.media.AMedia;
import org.zkoss.util.media.Media;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.IdSpace;
import org.zkoss.zk.ui.Path;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zk.ui.event.UploadEvent;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.Selectors;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.NRange;
import org.zkoss.zss.api.NRange.ApplyBorderType;
import org.zkoss.zss.api.NRange.DeleteShift;
import org.zkoss.zss.api.NRange.InsertCopyOrigin;
import org.zkoss.zss.api.NRange.InsertShift;
import org.zkoss.zss.api.NRanges;
import org.zkoss.zss.api.NSheetAnchor;
import org.zkoss.zss.api.model.NBook;
import org.zkoss.zss.api.model.NCellStyle;
import org.zkoss.zss.api.model.NCellStyle.BorderType;
import org.zkoss.zss.api.model.NChart.Grouping;
import org.zkoss.zss.api.model.NChart.LegendPosition;
import org.zkoss.zss.api.model.NChart;
import org.zkoss.zss.api.model.NChartData;
import org.zkoss.zss.api.model.NChartUtil;
import org.zkoss.zss.api.model.NFont.Boldweight;
import org.zkoss.zss.api.model.NFont.Underline;
import org.zkoss.zss.api.model.NPicture.Format;
import org.zkoss.zss.api.model.NPictureUtil;
import org.zkoss.zss.api.model.NSheet;
import org.zkoss.zss.api.ui.NSpreadsheet;
import org.zkoss.zss.essential.util.ClientUtil;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Combobox;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Menu;
import org.zkoss.zul.Menupopup;
import org.zkoss.zul.Toolbar;
import org.zkoss.zul.Toolbarbutton;
import org.zkoss.zul.impl.XulElement;

public class SheetbarCtrl extends SelectorComposer<Component> {
	private static final long serialVersionUID = 1L;

	private static final String ON_INIT_SPREADSHEET = "onInitSpreadsheet";

	String ssid;
	String sspath;
	NSpreadsheet nss;

	Toolbar sheetBar;
	
	@Wire
	Toolbarbutton addSheet;
	
	@Wire
	Menupopup sheetPopup;
	
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		sheetBar = (Toolbar)comp;
		Map arg = Executions.getCurrent().getArg();
		ssid = (String) arg.get("spreadsheetId");
		sspath = (String) arg.get("spreadsheetPath");

		comp.addEventListener(ON_INIT_SPREADSHEET, new EventListener<Event>() {
			public void onEvent(Event event) throws Exception {
				getSelf().removeEventListener(ON_INIT_SPREADSHEET, this);
				postInitSpreadsheet();
			}
		});
		Events.postEvent(new Event(ON_INIT_SPREADSHEET, comp));

		String script = "focusSS(this,'" + ssid + "')";
		for (Component c : Selectors.find(comp, "toolbarbutton")) {
			((XulElement) c).setWidgetListener("onClick", script);
		}
		for (Component c : Selectors.find(comp, "combobox")) {
			((XulElement) c).setWidgetListener("onClick", script);
			((XulElement) c).setWidgetListener("onSelect", script);
		}
		
		
	}

	public void postInitSpreadsheet() {
		Component comp = getSelf();

		Spreadsheet ss = (Spreadsheet) comp.getFellowIfAny(ssid);
		if (ss == null && sspath != null) {
			IdSpace o = comp.getSpaceOwner();
			if (o != null) {
				ss = (Spreadsheet) Path.getComponent(o, sspath);
			}
		}
		if (ss == null && sspath != null) {
			ss = (Spreadsheet) Path.getComponent(sspath);
		}

		if (ss == null) {
			throw new IllegalAddException(
					"spreadsheet component not found with id " + ssid);
		}

		nss = new NSpreadsheet(ss);
		ssid = null;

		
		
		refreshBar();
	}

	private void refreshBar() {
		NSheet sheet = nss.getSelectedSheet();
		NBook book = sheet.getBook();
		
		for(Component c : new ArrayList<Component>(sheetBar.getChildren())){
			if(c instanceof Toolbarbutton && !addSheet.equals(c)){
				sheetBar.removeChild(c);
			}
		}
		
		int size = book.getNumberOfSheets();
		for(int i=0;i<size;i++){
			NSheet s = book.getSheetAt(i);
			Toolbarbutton sheetBtn = new Toolbarbutton(s.getSheetName());
			sheetBtn.setMode("toggle");
			if(s.equals(sheet)){
				sheetBtn.setChecked(true);
			}
			sheetBtn.addEventListener(Events.ON_RIGHT_CLICK, new EventListener(){

				public void onEvent(Event event) throws Exception {
					Toolbarbutton sheetBtn = ((Toolbarbutton)event.getTarget());
					sheetPopup.open(sheetBtn);
					sheetPopup.setAttribute("targetSheet", sheetBtn.getLabel());
				}});
			sheetBtn.addEventListener(Events.ON_CHECK, new EventListener(){

				public void onEvent(Event event) throws Exception {
					Toolbarbutton sheetBtn = ((Toolbarbutton)event.getTarget()); 
					nss.setSelectedSheet(sheetBtn.getLabel());
					sheetBtn.setChecked(true);
					
					for(Component c : sheetBar.getChildren()){
						if(c instanceof Toolbarbutton && !addSheet.equals(c) && !c.equals(sheetBtn)){
							((Toolbarbutton)c).setChecked(false);
						}
					}
				}});
		
			sheetBar.appendChild(sheetBtn);
		}
		
	}
	private void showProtectionMessage(){
		ClientUtil.showWarn("Cann't modify a protected sheet/area");
	}
	
	
	// some test
	@Listen("onClick=#addSheet")
	public void onAddSheet() {
		NRange dest = NRanges.range(nss.getSelectedSheet());

		NSheet sheet = dest.createSheet(null);

		nss.setSelectedSheet(sheet.getSheetName());
		refreshBar();
	}

	@Listen("onClick=#deleteSheet")
	public void onDeleteSheet() {
		NSheet selectedSheet = nss.getSelectedSheet();
		NBook book = selectedSheet.getBook();
		
		String name = (String)sheetPopup.getAttribute("targetSheet");
		NSheet sheet = book.getSheet(name);
		
		NRange dest = NRanges.range(sheet);
		if (dest.isProtected()) {
			showProtectionMessage();
			return;
		}

		if (dest.getBook().getNumberOfSheets() <= 1) {
			ClientUtil.showWarn("Cann't delete last sheet");
			return;

		}

		int index = book.getSheetIndex(sheet);
		if (index != 0) {
			index--;
		}

		dest.deleteSheet();
		
		if(selectedSheet.equals(sheet)){
			// you have to handle the selected sheet manually as well.
			nss.setSelectedSheet(book.getSheetAt(index).getSheetName());
		}
		
		refreshBar();
	}
	
	
	@Listen("onClick=#renameSheet")
	public void onReanmeSheet() {
		NSheet selectedSheet = nss.getSelectedSheet();
		NBook book = selectedSheet.getBook();
		
		String name = (String)sheetPopup.getAttribute("targetSheet");
		NSheet sheet = book.getSheet(name);
		
		NRange dest = NRanges.range(sheet);
		if (dest.isProtected()) {
			showProtectionMessage();
			return;
		}

		
		//TODO , and rename dialog;
		StringBuilder sb = new StringBuilder();
		for(char c:name.toCharArray()){
			sb.insert(0, c);
		}
		
		String newname = sb.toString();
		
		if(book.getSheet(newname)!=null){
			ClientUtil.showWarn("Use another name");
			return;
		}
		
		
		dest.setSheetName(newname);
		refreshBar();
	}
	
	
	@Listen("onClick=#protectSheet")
	public void onProtectSheet() {
		NSheet selectedSheet = nss.getSelectedSheet();
		NBook book = selectedSheet.getBook();
		
		String name = (String)sheetPopup.getAttribute("targetSheet");
		NSheet sheet = book.getSheet(name);
		
		NRange dest = NRanges.range(sheet);
		
		//TODO, ask password is possible ..
//		if (dest.isProtected()) {
//			showProtectionMessage();
//			return;
//		}

		String newpassword = "1234";
		if(dest.isProtected()){
			dest.protectSheet(null);
		}else{
			dest.protectSheet(newpassword);
		}
		
		refreshBar();
	}
	
	@Listen("onClick=#moveSheetLeft")
	public void onMoveSheetLeft() {
		NSheet selectedSheet = nss.getSelectedSheet();
		NBook book = selectedSheet.getBook();
		
		String name = (String)sheetPopup.getAttribute("targetSheet");
		NSheet sheet = book.getSheet(name);
		
		NRange dest = NRanges.range(sheet);
		if (dest.isProtected()) {
			showProtectionMessage();
			return;
		}

		int index = book.getSheetIndex(sheet);
		if(index==0){
			return;
		}
		
		dest.setSheetOrder(index-1);
		refreshBar();
	}
	
	@Listen("onClick=#moveSheetRight")
	public void onMoveSheetRight() {
		NSheet selectedSheet = nss.getSelectedSheet();
		NBook book = selectedSheet.getBook();
		
		String name = (String)sheetPopup.getAttribute("targetSheet");
		NSheet sheet = book.getSheet(name);
		
		NRange dest = NRanges.range(sheet);
		if (dest.isProtected()) {
			showProtectionMessage();
			return;
		}

		int index = book.getSheetIndex(sheet);
		if(index==book.getNumberOfSheets()-1){
			return;
		}
		
		dest.setSheetOrder(index+1);
		refreshBar();
	}
}
