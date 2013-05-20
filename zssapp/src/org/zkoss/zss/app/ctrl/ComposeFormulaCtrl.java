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

//import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.HtmlBasedComponent;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.api.Range;
//import org.zkoss.zss.api.Range.CellType;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.app.formula.FormulaMetaInfo;
import org.zkoss.zss.app.zul.Dialog;
import org.zkoss.zss.app.zul.Zssapp;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zss.ui.event.CellSelectionEvent;
import org.zkoss.zul.Button;
import org.zkoss.zul.Label;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Listcell;
import org.zkoss.zul.Listitem;
import org.zkoss.zul.ListitemRenderer;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.SimpleListModel;
import org.zkoss.zul.Textbox;
import org.zkoss.zul.impl.InputElement;

/**
 * @author Sam
 *
 */
public class ComposeFormulaCtrl extends GenericForwardComposer {
	/* self dialog */
	private Dialog _composeFormulaDialog;
	private Label formulaStart;
	private Label formulaEnd;
	private Textbox composeFormulaTextbox;
	private Listbox argsListbox;
	private Label description;
	private Button okBtn;
	
	private FormulaMetaInfo info;
	private int focusToIndex = -1;
	private List<ArgWrapper> args;
	
	private List<InputElement> inputs = new LinkedList<InputElement>();
	
	/* whether move focus to next input component or not */
	private boolean movedToNext = false;
	
	/* current focus component */
	private InputElement focusComponent;

	private EventListener onCellSelected = new EventListener(){
		public void onEvent(Event event) throws Exception {
			CellSelectionEvent evt = (CellSelectionEvent) event;
			
			Range c = Ranges.range(evt.getSheet(), evt.getTop(), evt.getLeft());
//			Cell cell = Utils.getCell(evt.getSheet(), evt.getTop(), evt.getLeft());
			if (c.getCellData().getType() == CellType.BLANK) {
				c.setCellEditText("0");
			}
			if (focusComponent != null) {
				focusComponent.setText(getDesktopWorkbenchContext().getWorkbookCtrl().getCurrentCellPosition());
				Events.postEvent(Events.ON_CHANGE, focusComponent, null);
			}
		}};
	
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		formulaEnd.setValue(")");
		composeFormulaTextbox.addEventListener(Events.ON_CHANGE, new EventListener() {
			public void onEvent(Event event) throws Exception {
				decomposeFormula();
			}
		});
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
						movedToNext = false;
					}
				});
				tb.addEventListener(Events.ON_FOCUS, new EventListener() {
					public void onEvent(Event event) throws Exception {
						ArgWrapper last = args.get(args.size() - 1);
						if (last.equals(arg) && info.isMultipleParameter()) {
							focusToIndex = args.size() - 1;
							args.add(createNextArg());
							argsListbox.setModel(newListModelInstance(args));
						} else {
							focusComponent = tb;
						}
						
						if (movedToNext == true && focusComponent != null)
							focusComponent.focus();
						movedToNext = false;
					}
				});
				
				Listcell cell = new Listcell();
				cell.appendChild(tb);
				item.appendChild(cell);
			}

			@Override
			public void render(Listitem item, Object data, int index)
					throws Exception {
				render(item, data);
			}
		});
		argsListbox.addEventListener("onAfterRender", new EventListener() {
			public void onEvent(Event event) throws Exception {
				int focusIdx = -1;
				if (focusToIndex >= 0 && focusToIndex < inputs.size()) {
					focusIdx = focusToIndex;
				} else if (inputs.size() > 1) {
					focusIdx = 0;
				}
				if (focusIdx >= 0) {
					inputs.get(focusIdx).focus();
					focusComponent = inputs.get(focusIdx);
				}
			}
		});
	}
	
	public void onOpen$_composeFormulaDialog(ForwardEvent evt) {
		info = (FormulaMetaInfo) evt.getOrigin().getData();
		formulaStart.setValue("=" + info.getFunction() + "(");
		composeFormulaTextbox.setText(null);
		args = createArgs(info.getRequiredParameter(), info.getParameterNames());
		argsListbox.setModel(newListModelInstance(args));
		composeFormulaTextbox.focus();
		getDesktopWorkbenchContext().getWorkbookCtrl().addEventListener(org.zkoss.zss.ui.event.Events.ON_CELL_SELECTION, onCellSelected);
	}
	
	public void onClose$_composeFormulaDialog() {
		getDesktopWorkbenchContext().getWorkbookCtrl().removeEventListener(org.zkoss.zss.ui.event.Events.ON_CELL_SELECTION, onCellSelected);
	}
	
	private void init() {
		formulaStart.setValue("=" + info.getFunction() + "(");
		description.setValue(info.getDescription());
		
		args = createArgs(info.getRequiredParameter(), info.getParameterNames());
		argsListbox.setModel(newListModelInstance(args));
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
				movedToNext = true;
				break;
			}
		}
	}
	
	private void composeFormula() {
		StringBuilder strBuilder = new StringBuilder();
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
		composeFormulaTextbox.setText(strBuilder.toString());
	}

	private ArgWrapper createNextArg() {
		Integer num = null;
		String argName = info.getMultipleParameter();
		ArgWrapper last = args.get(args.size() - 1);
		try {
			num = Integer.parseInt(last.getName().replace(argName, ""));
			num++;
		} catch (NumberFormatException ex) {
		}
		return new ArgWrapper(args.size(), info.getMultipleParameter() + (num != null ? num : ""), "");
	}
	
	private void decomposeFormula() {
		String input = composeFormulaTextbox.getText();
		String[] arg = input.split(",");
		if (arg.length > args.size())
			if (info.isMultipleParameter()) {
				int diff = arg.length - args.size();
				for (int i = 0; i < diff; i++) {
					args.add(createNextArg());
				}
			} else {
				Messagebox.show("You've entered too many arguments for this function");
				return;
			}
		for (int i = 0; i < args.size() && i < arg.length; i++) {
			String val = arg[i].trim();
			args.get(i).setValue(val);
		}
		argsListbox.setModel(newListModelInstance(args));
	}
	
	public void onClick$okBtn() {
		getDesktopWorkbenchContext().getWorkbookCtrl().insertFormula(info.getRowIndex(), info.getColIndex(), formulaStart.getValue() + composeFormulaTextbox.getText() + formulaEnd.getValue());
		//Note. insert formula may throw exception and won't fire book content changed event, need to fire own event to update UI
		getDesktopWorkbenchContext().fireContentsChanged();
		_composeFormulaDialog.fireOnClose(null);
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
		return Zssapp.getDesktopWorkbenchContext(self);
	}
}