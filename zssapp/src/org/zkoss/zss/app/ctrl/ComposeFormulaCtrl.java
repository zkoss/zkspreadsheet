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

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zk.ui.event.InputEvent;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.app.formula.FormulaMetaInfo;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zul.Button;
import org.zkoss.zul.Label;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Listcell;
import org.zkoss.zul.Listitem;
import org.zkoss.zul.ListitemRenderer;
import org.zkoss.zul.SimpleListModel;
import org.zkoss.zul.Textbox;

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
	private List<ArgWrapper> args;

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
		argsListbox.setModel(new SimpleListModel(args));
		argsListbox.setItemRenderer(new ListitemRenderer() {
			public void render(Listitem item, Object data) throws Exception {
				final ArgWrapper arg = (ArgWrapper)data;
				item.setValue(arg);
				item.appendChild(new Listcell(arg.getName()));
				final Textbox tb = new Textbox(arg.getValue());
				tb.addEventListener(Events.ON_CHANGE, new EventListener() {
					public void onEvent(Event event) throws Exception {
						InputEvent evt = (InputEvent) event;
						arg.setValue(evt.getValue());
						composeFormula();
					}
				});
				
				Listcell cell = new Listcell();
				cell.appendChild(tb);
				item.appendChild(cell);
			}
		});
	}
	private void composeFormula() {
		String result = "=" + info.getFunction() + "(";
		boolean first = true;
		for (ArgWrapper w : args) {
			String arg = w.getValue();
			if (!"".equals(arg)) {
				if (first)
					first = false;
				else if (!first)
					result += ",";
				result += arg;
			}				
		}
		result += ")";
		composeFormulaTextbox.setText(result);
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
		argsListbox.setModel(new SimpleListModel(args));
	}
	
	public void onClick$okBtn() {
		//DesktopWorkbenchContext.getInstance(desktop).
		//	getWorkbookCtrl().insertFormula();
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
}