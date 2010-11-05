package org.zkoss.zss.app;

import java.util.HashMap;
import java.util.Map;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.zk.ui.Path;
//import org.zkoss.zss.model.Sheet;
//import org.zkoss.zss.model.TextHAlign;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.Utils;
//import org.zkoss.zss.ui.impl.Styles;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Window;


public class FormulaCategoryHelper {
	Spreadsheet spreadsheet;
	Map mapLabelListbox;

	public FormulaCategoryHelper(Spreadsheet spreadsheet) {
		this.spreadsheet = spreadsheet;
		mapFormatValues();
	}

	public void mapFormatValues() {
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

	public void onOK() {
		try {
			Listbox fc_category = (Listbox) Path.getComponent("//p1/mainWin/formulaCategory/fc_category");
			if (fc_category.getSelectedItem() == null) {
				return;
			}
			
			Listbox selectedList = (Listbox) Path.getComponent("//p1/mainWin/formulaCategory/" + mapLabelListbox.get(fc_category.getSelectedItem().getLabel()));

			String formula = "=" + selectedList.getSelectedItem().getLabel().toString()	+ "()";

			int left = spreadsheet.getSelection().getLeft();
			int top = spreadsheet.getSelection().getTop();
			Cell cell = Utils.getOrCreateCell(spreadsheet.getSelectedSheet(), top, left);
			cell.setCellFormula(formula);
		} catch (Exception e) {
			e.printStackTrace();
		}

		//Window win = (Window) Path.getComponent("//p1/formulaCategoryWin");
		//win.detach();
	}
}
