package org.zkoss.zss.app;

import java.util.HashMap;
import java.util.Map;

//import org.zkoss.poi.ss.usermodel.CellStyle;
//import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.zk.ui.Path;
import org.zkoss.zk.ui.UiException;
//import org.zkoss.zss.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;
//import org.zkoss.zss.ui.impl.Styles;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Listitem;
import org.zkoss.zul.Window;


public class FormatNumberHelper {
	Spreadsheet spreadsheet;
	Map mapLabelListbox;

	public FormatNumberHelper(Spreadsheet spreadsheet) {
		this.spreadsheet = spreadsheet;
		mapFormatValues();
	}

	public void mapFormatValues() {
		mapLabelListbox = new HashMap();

		mapLabelListbox.put("General", "mfn_general");
		mapLabelListbox.put("Number", "mfn_number");
		mapLabelListbox.put("Currency", "mfn_currency");
		mapLabelListbox.put("Accounting", "mfn_accounting");
		mapLabelListbox.put("Date", "mfn_date");
		mapLabelListbox.put("Time", "mfn_time");
		mapLabelListbox.put("Percentage", "mfn_percentage");
		mapLabelListbox.put("Fraction", "mfn_fraction");
		mapLabelListbox.put("Scientific", "mfn_scientific");
		mapLabelListbox.put("Text", "mfn_text");
		mapLabelListbox.put("Special", "mfn_special");
	}

	public void onOK() {
//TODO remove me and comments
throw new UiException("format setting is not implmented yet");		
/* 
		try {
			Listbox mfn_category = (Listbox) Path.getComponent("//p1/mainWin/formatNumberWin/mfn_category");
			if (mfn_category.getSelectedItem() == null) {
				closeFormatNumberWindow();
				return;
			}
			
//TODO undo/redo			
//			spreadsheet.pushCellState();
			
			Listbox selectedList = (Listbox) Path.getComponent("//p1/mainWin/formatNumberWin/" + mapLabelListbox.get(mfn_category.getSelectedItem().getLabel()));
			Listitem selectedItem = selectedList.getSelectedItem();

			if (selectedItem != null) {
				String formatCodes = selectedItem.getValue().toString();

				// can't set default selected item for combobox, it's just a text.

				int left = spreadsheet.getSelection().getLeft();
				int right = spreadsheet.getSelection().getRight();
				int top = spreadsheet.getSelection().getTop();
				int bottom = spreadsheet.getSelection().getBottom();
				Sheet sheet = spreadsheet.getSelectedSheet();
				for (int row = top; row <= bottom; row++) {
					for (int col = left; col <= right; col++) {
						Styles.setFormatCodes(sheet, row, col, formatCodes);
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("wh FormatNumberHelper null pointer");
		}

		closeFormatNumberWindow();
*/		
	}
	
	private void closeFormatNumberWindow() {
		Window formatNumberWin = (Window) Path.getComponent("//p1/mainWin/formatNumberWin");
		formatNumberWin.setVisible(false);
		spreadsheet.setHighlight(null);
	}
}
