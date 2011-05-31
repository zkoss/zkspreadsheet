/* InsertFormulaCtrl2.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 25, 2010 5:54:06 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.ctrl;

import java.util.Collections;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zkplus.databind.BindingListModelList;
import org.zkoss.zss.app.formula.FormulaMetaInfo;
import org.zkoss.zss.app.formula.Formulas;
import org.zkoss.zss.app.zul.Dialog;
import org.zkoss.zss.app.zul.Zssapp;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zss.ui.Position;
import org.zkoss.zul.Button;
import org.zkoss.zul.Combobox;
import org.zkoss.zul.Label;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Listitem;
import org.zkoss.zul.ListitemRenderer;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.SimpleListModel;
import org.zkoss.zul.Textbox;
import org.zkoss.zul.Window;

/**
 * @author Sam
 *
 */
public class InsertFormulaCtrl2 extends GenericForwardComposer {
	
	private final static String ALL = "All";
	
	private Dialog _insertFormulaDialog;
	private Textbox searchTextbox;
	private Button searchBtn;
	
	private Combobox categoryCombobox;
	private Listbox functionListbox;
	private Label expression;
	private Label description;
	
	private Button okBtn;
	
	LinkedHashMap<String, List<FormulaMetaInfo>> formulaInfos;
	
	private int rowIdx;
	private int colIdx;
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		formulaInfos = Formulas.getFormulaInfos();
		
		LinkedList<String> categoryArys = new LinkedList<String>();
		categoryArys.add(ALL);
		categoryArys.addAll(formulaInfos.keySet());
		categoryCombobox.setModel(
				new BindingListModelList(categoryArys, false));
		categoryCombobox.addEventListener("onAfterRender", new EventListener() {
			public void onEvent(Event event) throws Exception {
				categoryCombobox.setSelectedIndex(0);
				initFunctionListbox();
			}
		});
		
		functionListbox.setItemRenderer(new ListitemRenderer() {
			public void render(Listitem item, Object data) throws Exception {
				FormulaMetaInfo info = (FormulaMetaInfo)data;
				item.setLabel(info.getFunction());
				item.setValue(info);
			}
		});
		functionListbox.addEventListener(Events.ON_DOUBLE_CLICK, new EventListener() {
			public void onEvent(Event event) throws Exception {
				openComposeFormulaDialog();
			}
		});
	}
	
	public void onOpen$_insertFormulaDialog() {
		try {
			_insertFormulaDialog.setMode(Window.MODAL);
		} catch (InterruptedException e) {
		}
		searchTextbox.setText(null);
		initFunctionListbox();
		searchTextbox.focus();
		
		Position pos = getDesktopWorkbenchContext().getWorkbookCtrl().getCellFocus();
		rowIdx = pos.getRow();
		colIdx = pos.getColumn();
	}
	
	public void onSelect$categoryCombobox() {
		initFunctionListbox();
	}
	
	private void initFunctionListbox() {
		String category = categoryCombobox.getText();
		if (category == null)
			return;
		
		List<FormulaMetaInfo> ary = formulaInfos.get(category);
		if (ary == null) {
			//means all
			ary = new LinkedList<FormulaMetaInfo>();
			for (List<FormulaMetaInfo> infoAry : formulaInfos.values()) {
				ary.addAll(infoAry);
			}
		}
		functionListbox.setModel(new SimpleListModel(ary));
	}
	
	public void onClick$functionListbox() {
		FormulaMetaInfo info = (FormulaMetaInfo)functionListbox.getSelectedItem().getValue();
		expression.setValue(info.getExpression());
		description.setValue(info.getDescription());
	}
	
	public void onClick$okBtn() {
		openComposeFormulaDialog();
	}
	
	private void openComposeFormulaDialog() {
		Listitem item = (Listitem)functionListbox.getSelectedItem();
		if (item == null) {
			try {
				Messagebox.show("Select a function");
				return;
			} catch (InterruptedException e) {
			}
		}	
		
		FormulaMetaInfo info = (FormulaMetaInfo) item.getValue();
		if (info.getRequiredParameter() == 0) {
			getDesktopWorkbenchContext().getWorkbookCtrl().insertFormula(rowIdx, colIdx, "=" + info.getFunction() + "()");
		} else {
			info.setRowIndex(rowIdx);
			info.setColIndex(colIdx);
			getDesktopWorkbenchContext().getWorkbenchCtrl().
				openComposeFormulaDialog((FormulaMetaInfo)item.getValue());
		}

		_insertFormulaDialog.fireOnClose(null);
	}
	
	//TODO: shall I use echo event to showbusy, need to test speed 
	private List<FormulaMetaInfo> search(String searchFor) {
		searchFor = searchFor.toLowerCase();
		LinkedList<FormulaMetaInfo> searchResult = new LinkedList<FormulaMetaInfo>();
		for (List<FormulaMetaInfo> infoAry : formulaInfos.values()) {
			for (FormulaMetaInfo info : infoAry) {
				if (info.getFunction().toLowerCase().contains(searchFor) ||
					info.getExpression().toLowerCase().contains(searchFor) ||
					info.getDescription().toLowerCase().contains(searchFor))
					searchResult.add(info);
			}
		}
		Collections.sort(searchResult, new Comparator<FormulaMetaInfo>() {
			public int compare(FormulaMetaInfo o1, FormulaMetaInfo o2) {
				return o1.getFunction().compareTo(o2.getFunction());
			}
		});
		return searchResult;
	}
	public void onOK$searchTextbox() {
		String searchFor = searchTextbox.getValue();
		if (searchFor == null || "".equals(searchFor))
			return;
		functionListbox.setModel(new SimpleListModel(search(searchFor)));
	}
	
	public void onClick$searchBtn() {
		String searchFor = searchTextbox.getText();
		if (searchFor == null || "".equals(searchFor))
			return;
		functionListbox.setModel(new SimpleListModel(search(searchFor)));
	}
	protected DesktopWorkbenchContext getDesktopWorkbenchContext() {
		return Zssapp.getDesktopWorkbenchContext(self);
	}
}