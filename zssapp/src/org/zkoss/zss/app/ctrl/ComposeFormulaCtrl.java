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

import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.app.formula.FormulaMetaInfo;
import org.zkoss.zul.Label;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Textbox;

/**
 * @author Sam
 *
 */
public class ComposeFormulaCtrl extends GenericForwardComposer {

	private Textbox composeFormulaTextbox;
	private Listbox argsListbox;
	private Label description;
	
	private FormulaMetaInfo info;
	private HashMap<Integer, String> args;

	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		info  = (FormulaMetaInfo)Executions.getCurrent().getArg().get(Consts.KEY_ARG_FORMULA_METAINFO);
		composeFormulaTextbox.setText(info.getFunction());
		
		description.setValue(info.getDescription());
	}
	
//	private List<ArgWrapper> createArgs(FormulaMetaInfo info) {
//		LinkedList<ArgWrapper> ary = new LinkedList<ArgWrapper>();
//		for (int i = 0; i < info.getRequiredParameter(); i++) {
//			//ary.add(new ArgWrapper(i, value))
//		}
//	}
	
	private class ArgWrapper {
		int index;
		String value;
		//type
		public ArgWrapper(int index, String value) {
			super();
			this.index = index;
			this.value = value;
		}
		public int getIndex() {
			return index;
		}
		public void setIndex(int index) {
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