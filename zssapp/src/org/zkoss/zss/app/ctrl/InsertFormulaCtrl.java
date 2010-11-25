/* InsertFormulaCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 24, 2010 5:02:37 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.ctrl;

import java.util.HashMap;
import java.util.Map;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zul.Button;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Listitem;
import org.zkoss.zul.Window;

/**
 * @author Sam
 *
 */
public class InsertFormulaCtrl extends GenericForwardComposer {

	
	Map<String, String> mapLabelListbox;
	
	private Window formulaWin;
	
	private Listbox fc_category;
	
	private Button insertFormulaBtn;
	
	public InsertFormulaCtrl() {
		mapFormatValues();
	}
	private void mapFormatValues() {
		mapLabelListbox = new HashMap();

		mapLabelListbox.put("General", "fc_general");
		mapLabelListbox.put("Math", "fc_math");
		mapLabelListbox.put("Logical", "fc_logical");
		mapLabelListbox.put("Statistical", "fc_statistical");
		mapLabelListbox.put("Date and Time", "fc_dateAndTime");
		mapLabelListbox.put("Information", "fc_information");
		mapLabelListbox.put("Text and Data", "fc_textAndData");
		mapLabelListbox.put("Financial", "fc_financial");
		mapLabelListbox.put("Engineering", "fc_engineering");
	}

	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		fc_category.setSelectedIndex(0);
	}

	public void onClick$insertFormulaBtn() {
		Listitem seld = fc_category.getSelectedItem();
		if (seld == null)
			return;
		
		Listbox seldListbox = (Listbox)formulaWin.getFellow(
				mapLabelListbox.get(fc_category.getSelectedItem().getLabel()) );
		
		Listitem seldName =  seldListbox.getSelectedItem();
		if (seldName == null)
			return;
		
		String formula = "=" + seldName.getLabel().toString()	+ "()";
		DesktopWorkbenchContext.getInstance(desktop).insertFormula(formula);
		formulaWin.setVisible(false);
	}
}
