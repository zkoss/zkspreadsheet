package org.zkoss.zss.app;

//import org.zkoss.poi.ss.usermodel.Cell;
//import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.zk.ui.Path;
//import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.CellOperationUtil.CellStyleApplier;
import org.zkoss.zss.api.CellVisitor;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.EditableCellStyle;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zul.Checkbox;
import org.zkoss.zul.Window;

public class FormatProtectionHelper {
	Spreadsheet spreadsheet;

	public FormatProtectionHelper(Spreadsheet spreadsheet) {
		this.spreadsheet = spreadsheet;
	}

	public void onOK() {
		try {
			final Checkbox mfp_locked = (Checkbox) Path.getComponent("//p1/mainWin/formatMainWin/formatProtection/mfp_locked");
			final Checkbox mfp_hidden = (Checkbox) Path.getComponent("//p1/mainWin/formatMainWin/formatProtection/mfp_hidden");
//
//			int left = spreadsheet.getSelection().getLeft();
//			int right = spreadsheet.getSelection().getRight();
//			int top = spreadsheet.getSelection().getTop();
//			int bottom = spreadsheet.getSelection().getBottom();
//			Sheet sheet = spreadsheet.getSelectedSheet();
			
			Range range = Ranges.range(spreadsheet.getSelectedSheet(),spreadsheet.getSelection());
			
			
			CellOperationUtil.applyCellStyle(range, new CellStyleApplier() {
				
				@Override
				public boolean ignore(Range cellRange, CellStyle oldCellstyle) {
					return oldCellstyle.isHidden()==mfp_hidden.isChecked() && oldCellstyle.isLocked()==mfp_locked.isChecked();
				}
				
				@Override
				public void apply(Range cellRange, EditableCellStyle newCellstyle) {
					newCellstyle.setHidden(mfp_hidden.isChecked());
					newCellstyle.setLocked(mfp_hidden.isChecked());
				}
			});
//			
//			for (int row = top; row <= bottom; row++) {
//				for (int col = left; col <= right; col++) {
//					final Cell cell = Utils.getCell(sheet, row, col);
//					if (cell != null) {
//						CellStyle cs = cell.getCellStyle();
//						cs.setHidden(mfp_hidden.isChecked());
//						cs.setLocked(mfp_hidden.isChecked());
//						//Styles.setLocked(sheet, row, col, mfp_locked.isChecked());
//						//Styles.setHidden(sheet, row, col, mfp_hidden.isChecked());
//					}
//				}
//			}
		} catch (Exception e) {
			e.printStackTrace();
//			System.out.println("wh FormatNumberHelper null pointer");
		}

		Window formatMainWin = (Window) Path.getComponent("//p1/mainWin/formatMainWin");
		formatMainWin.setVisible(false);
		spreadsheet.setHighlight(null);
	}
}
