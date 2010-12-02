/* ComposeFormulaDialog.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 26, 2010 10:41:33 AM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.ctrl;

import java.util.LinkedList;
import java.util.List;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.HtmlBasedComponent;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.app.formula.FormulaMetaInfo;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.ui.event.CellSelectionEvent;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zul.Button;
import org.zkoss.zul.Label;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Listcell;
import org.zkoss.zul.Listitem;
import org.zkoss.zul.ListitemRenderer;
import org.zkoss.zul.SimpleListModel;
import org.zkoss.zul.Textbox;
import org.zkoss.zul.impl.api.InputElement;

/**
 * @author Sam
 *
 */
public class ComposeFormulaCtrl extends GenericForwardComposer {

	private Textbox composeFormulaTextbox;
	private Listbox argsListbox;
	private Label description;
	private Button okBtn;
	
	private FormulaMetaInfo info;
	private int focusToIndex = -1;
	private List<ArgWrapper> args;
	
	private List<InputElement> inputs = new LinkedList<InputElement>();
	private InputElement focusComponent;

	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		info  = (FormulaMetaInfo)Executions.getCurrent().getArg().get(Consts.KEY_ARG_FORMULA_METAINFO);
		composeFormulaTextbox.setText("=" + info.getFunction() + "()");
		composeFormulaTextbox.addEventListener(Events.ON_CHANGE, new EventListener() {
			public void onEvent(Event event) throws Exception {
				decomposeFormula();
			}
		});
		
		description.setValue(info.getDescription());
		
		args = createArgs(info.getRequiredParameter(), info.getParameterNames());
		argsListbox.setModel(newListModelInstance(args));
		argsListbox.setItemRenderer(new ListitemRenderer() {
			public void render(Listitem item, Object data) throws Exception {
				final ArgWrapper arg = (ArgWrapper)data;
				item.setValue(arg);
				item.appendChild(new Listcell(arg.getName()));
				final Textbox tb = new Textbox(arg.getValue());
				inputs.add(tb);
				
				tb.addEventListener(Events.ON_CHANGE, new EventListener() {
					public void onEvent(Event event) throws Exception {
						arg.setValue(tb.getValue());
						composeFormula();
						moveFocusToNext(tb);
					}
				});
				tb.addEventListener(Events.ON_FOCUS, new EventListener() {
					public void onEvent(Event event) throws Exception {
						ArgWrapper last = args.get(args.size() - 1);
						if (last.equals(arg)) {
							String argName = info.getMultipleParameter();
							Integer num = null;
							try {
							 num = Integer.parseInt(last.getName().replace(argName, ""));
							 num++;
							} catch (NumberFormatException ex) {
							}
							focusToIndex = last.getIndex();
							args.add(new ArgWrapper(args.size(), info.getMultipleParameter() + (num != null ? num : ""), ""));
							argsListbox.setModel(newListModelInstance(args));
						} else
							focusComponent = tb;
					}
				});
				
				Listcell cell = new Listcell();
				cell.appendChild(tb);
				item.appendChild(cell);
			}
		});
		argsListbox.addEventListener("onAfterRender", new EventListener() {
			public void onEvent(Event event) throws Exception {
				int focusIdx = -1;
				System.out.println("focus to: " + focusToIndex);
				if (focusToIndex >= 0 && focusToIndex < inputs.size()) {
					focusIdx = focusToIndex;
				}
				else if (inputs.size() > 1) {
					focusIdx = 0;
				}
				System.out.println("set focu to " + focusIdx);
				if (focusIdx >= 0) {
					inputs.get(focusIdx).focus();
					focusComponent = inputs.get(focusIdx);
				}
			}
		});
		
		getDesktopWorkbenchContext().getWorkbookCtrl().addEventListener(org.zkoss.zss.ui.event.Events.ON_CELL_SELECTION, new EventListener() {
			public void onEvent(Event event) throws Exception {
				CellSelectionEvent evt = (CellSelectionEvent) event;
				Cell cell = Utils.getCell(evt.getSheet(), evt.getTop(), evt.getLeft());
				if (cell == null) {
					Ranges.range(evt.getSheet(), evt.getTop(), evt.getLeft()).setEditText("0");
				}
				if (focusComponent != null) {
					focusComponent.setText(getDesktopWorkbenchContext().getWorkbookCtrl().getCurrentCellPosition());
					Events.postEvent(Events.ON_CHANGE, focusComponent, null);
				}
			}
		});
	}

	private SimpleListModel newListModelInstance(List<ArgWrapper> ary) {
		inputs.clear();
		focusComponent = null;
		return new SimpleListModel(ary);
	}
	
	private void moveFocusToNext(HtmlBasedComponent current) {
		for (int i = 0; i < inputs.size(); i++) {
			InputElement c = inputs.get(i);
			if (c == current && (i + 1) < inputs.size()) {
				InputElement next = inputs.get(i + 1);
				next.focus();
				focusComponent = next;
				break;
			}
		}
	}
	
	private void composeFormula() {
		StringBuilder strBuilder = new StringBuilder();
		strBuilder.append("=" + info.getFunction() + "(");
		boolean first = true;
		for (int i = 0; i < args.size(); i++) {
			ArgWrapper curArg = args.get(i);
			String arg = curArg.getValue();
			if (first)
				first = false;
			else if (!first) {
				if (!"".equals(arg))
					strBuilder.append(",");
				else {
					for (int j = i + 1; j < args.size(); j++) {
						String val = args.get(j).getValue();
						if (!"".equals(val)) {
							strBuilder.append(",");
							break;
						}
					}
				}
			}
			if (!"".equals(arg)) {
				strBuilder.append(arg);
			}
		}
		strBuilder.append(")");
		composeFormulaTextbox.setText(strBuilder.toString());
	}
	
	private void decomposeFormula() {
		String input = composeFormulaTextbox.getText();
		int startIdx = -1, endIdx = -1;
		if (input == null || 
			((startIdx = input.indexOf("(")) == -1) || 
			((endIdx = input.lastIndexOf(")")) == -1)
			)
			return;
		startIdx += 1;
		input = input.substring(startIdx, endIdx);
		String[] arg = input.split(",");
		for (int i = 0; i < arg.length; i++) {
			String val = arg[i].trim();
			args.get(i).setValue(val);
		}
		argsListbox.setModel(newListModelInstance(args));
	}
	
	public void onClick$okBtn() {
		getDesktopWorkbenchContext().getWorkbookCtrl().insertFormula(composeFormulaTextbox.getText());
		//Note. insert formula may throw exception and won't fire book content changed event, need to fire own event to update UI
		getDesktopWorkbenchContext().fireContentsChanged();
		self.detach();
	}
	
	private List<ArgWrapper> createArgs(int numArg, String[] argNames) {
		LinkedList<ArgWrapper> ary = new LinkedList<ArgWrapper>();
		for (int i = 0; i < numArg; i++) {
			ary.add(new ArgWrapper(i, argNames[i], ""));
		}
		return ary;
	}
	
	private class ArgWrapper {
		Integer index;
		String name;
		String value;
		//String type
		
		public ArgWrapper(Integer index, String name, String value) {
			super();
			this.index = index;
			this.name = name;
			this.value = value;
		}
		public String getName() {
			return name;
		}
		public void setName(String name) {
			this.name = name;
		}
		public Integer getIndex() {
			return index;
		}
		public void setIndex(Integer index) {
			this.index = index;
		}
		public String getValue() {
			return value;
		}
		public void setValue(String value) {
			this.value = value;
		}
	}
	
	protected DesktopWorkbenchContext getDesktopWorkbenchContext() {
		return DesktopWorkbenchContext.getInstance(desktop);
	}
}