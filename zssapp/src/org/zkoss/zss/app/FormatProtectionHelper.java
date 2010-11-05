package org.zkoss.zss.app;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.zk.ui.Path;
//import org.zkoss.zss.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.Utils;
//import org.zkoss.zss.ui.impl.Styles;
import org.zkoss.zul.Checkbox;
import org.zkoss.zul.Window;

public class FormatProtectionHelper {
	Spreadsheet spreadsheet;

	public FormatProtectionHelper(Spreadsheet spreadsheet) {
		this.spreadsheet = spreadsheet;
	}

	public void onOK() {
		try {
			Checkbox mfp_locked = (Checkbox) Path.getComponent("//p1/mainWin/formatMainWin/formatProtection/mfp_locked");
			Checkbox mfp_hidden = (Checkbox) Path.getComponent("//p1/mainWin/formatMainWin/formatProtection/mfp_hidden");

			int left = spreadsheet.getSelection().getLeft();
			int right = spreadsheet.getSelection().getRight();
			int top = spreadsheet.getSelection().getTop();
			int bottom = spreadsheet.getSelection().getBottom();
			Sheet sheet = spreadsheet.getSelectedSheet();
			for (int row = top; row <= bottom; row++) {
				for (int col = left; col <= right; col++) {
					final Cell cell = Utils.getCell(sheet, row, col);
					if (cell != null) {
						CellStyle cs = cell.getCellStyle();
						cs.setHidden(mfp_hidden.isChecked());
						cs.setLocked(mfp_hidden.isChecked());
						//Styles.setLocked(sheet, row, col, mfp_locked.isChecked());
						//Styles.setHidden(sheet, row, col, mfp_hidden.isChecked());
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("wh FormatNumberHelper null pointer");
		}

		Window formatMainWin = (Window) Path.getComponent("//p1/mainWin/formatMainWin");
		formatMainWin.setVisible(false);
		spreadsheet.setHighlight(null);
	}
}
