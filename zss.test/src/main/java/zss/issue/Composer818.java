/**
 * This is about DemoComposer
 * @author mqiu
 * @date 10/27/2014
 * 
 * Description: a demo to show the formula issue
 * If a cell's formula references another formula cell and the second cell is
 * not preloaded, the formula value would be incorrect. The second cell would not 
 * calculate even after loaded.
 */

package zss.issue;

import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.CellData.CellType;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Window;

public class Composer818 extends SelectorComposer<Window> {

	private static final long serialVersionUID = -7360829196117880724L;

	private static final int START_ROW = 0;
	private static final int TITLE_END_ROW = 1;
	private static final int CONTENT_ROW = 2;
	private static final int START_COLUMN = 0;
	private static final int LAST_COLUMN = START_COLUMN + 20;

	private static final int CONTENT_NUM = 500;

	private int lastRow = CONTENT_ROW + CONTENT_NUM;
	private int lastColumn = Composer818.LAST_COLUMN;

	private int clickTimes = 0;

	@Wire
	Spreadsheet spreadsheet;

	Sheet srcSheet;
	Sheet sheet;

	public void doAfterCompose(Window comp) throws Exception {
		super.doAfterCompose(comp);
		srcSheet = spreadsheet.getSelectedSheet();
		sheet = Ranges.range(srcSheet).createSheet("Demo");
		spreadsheet.setSelectedSheet("Demo");

		CellRegion titleRegion = new CellRegion(START_ROW, START_COLUMN, TITLE_END_ROW, LAST_COLUMN);
		copyTitle(titleRegion, titleRegion);
		copyRow(CONTENT_ROW, lastRow);
		copyRowHeight(START_ROW, lastRow);
		copyColumnWidth(START_COLUMN, LAST_COLUMN);
		// preloading all the rows is a a workaround. But this would result in bad performance.
		// spreadsheet.setPreloadRowSize(lastRow + 1);

		for (int i = CONTENT_ROW; i < lastRow + 1; i++) {
			int realContentRow = i - CONTENT_ROW;

			// the first row is the sum-all row, and references all the below rows
			// this causes the issue
			if (realContentRow == 0) {
				Ranges.range(sheet, i, 2).getCellData().setValue("SUM-ALL");
				for (int j = 3; j < LAST_COLUMN + 1; j++) {
					CellRegion c = new CellRegion(i + 1, j, lastRow - 1, j);
					Ranges.range(sheet, i, j).getCellData().setValue("=sum(" + c.getReferenceString() + ")");
				}
			} else {
				CellRegion c = new CellRegion(i, 4, i, LAST_COLUMN);
				Ranges.range(sheet, i, 3).getCellData().setValue("=sum(" + c.getReferenceString() + ")");
			}
		}

		Ranges.range(srcSheet).deleteSheet();

		spreadsheet.setMaxVisibleColumns(lastColumn + 1);
		spreadsheet.setMaxVisibleRows(lastRow + 1);
	}

	@Listen("onCellClick=#spreadsheet")
	public void onCellClick() {

		// simulate populating data
		if (clickTimes == 0) {
			clickTimes++;
			for (int i = CONTENT_ROW + 1; i < lastRow + 1; i++) {
				Range cell = Ranges.range(sheet, i, 5);
				cell.setAutoRefresh(false);
				if (!cell.getCellData().getType().equals(CellType.FORMULA)) {
					cell.setCellValue(2);
				}
				//workaround
//				Range d = Ranges.range(sheet, i + 1, 3);
//				d.setAutoRefresh(false);
//				d.setCellValue(d.getCellValue());
			}
			Ranges.range(sheet).notifyChange();
		}
	}

	private void copyTitle(CellRegion srcRegion, CellRegion destRegion) {
		copyRange(srcRegion, destRegion);
	}

	private void copyRow(int srcRow, int destRow) {
		CellRegion srcRegion = new CellRegion(srcRow, START_COLUMN, srcRow, LAST_COLUMN);
		CellRegion destRegion = new CellRegion(srcRow, START_COLUMN, destRow, LAST_COLUMN);

		copyRange(srcRegion, destRegion);
	}

	private void copyRange(CellRegion srcRegion, CellRegion destRegion) {
		Range srcRange = Ranges.range(srcSheet, srcRegion.getReferenceString());
		Range destRange = Ranges.range(sheet, destRegion.getReferenceString());
		srcRange.pasteSpecial(destRange, Range.PasteType.ALL, Range.PasteOperation.ADD, false, false);
	}

	private void copyColumnWidth(int startColumn, int lastColumn) {
		for (int i = startColumn; i < lastColumn + 1; i++) {
			Range destRange = Ranges.range(sheet, 0, i).toColumnRange();
			destRange.setColumnWidth(srcSheet.getColumnWidth(i));
		}
	}

	private void copyRowHeight(int startRow, int lastRow) {
		for (int i = startRow; i < lastRow + 1; i++) {
			Range destRange = Ranges.range(sheet, i, 0).toRowRange();
			destRange.setRowHeight(srcSheet.getRowHeight(i), true);
		}
	}
}
