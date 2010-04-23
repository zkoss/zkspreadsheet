package org.zkoss.zss.app;

import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.zkoss.zk.ui.Path;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Range;
//import org.zkoss.zss.model.Sheet;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Listitem;
import org.zkoss.zul.Textbox;
import org.zkoss.zul.Window;

public class RangeHelper {

	Spreadsheet spreadsheet;

	public RangeHelper(Spreadsheet spreadsheet) {
		this.spreadsheet = spreadsheet;
	}

	public void dispatcher(String type) {
		if (type.equals("add")) {
			onRangeAdd();
		} else if (type.equals("delete")) {
			onRangeDelete();
		} else if (type.equals("choose")) {
			onRangeChoose();
		} else if (type.equals("selectChangeAdd")) {
			onRangeSelectChange("add");
		} else if (type.equals("selectChangeDelete")) {
			onRangeSelectChange("delete");
		} else if (type.equals("selectChangeChoose")) {
			onRangeSelectChange("choose");
		}

	}

	// dummy
	public String[] getRangeList() {
		Book book = spreadsheet.getBook();
		int len = book.getNumberOfNames();
		String[] names = new String[len];
		for(int j=0; j < len; ++j) {
			names[j] = book.getNameAt(j).getNameName();
		}
		return names;
	}

	// dummy
	public void onRangeAdd() {
		Book book = spreadsheet.getBook();
		Sheet sheet = spreadsheet.getSelectedSheet();
		int sheetindex = book.getSheetIndex(sheet);
		int left = spreadsheet.getSelection().getLeft();
		int top = spreadsheet.getSelection().getTop();
		int right = spreadsheet.getSelection().getRight();
		int bottom = spreadsheet.getSelection().getBottom();
		Textbox fsaw_filename = (Textbox) Path
				.getComponent("//p1/rangeAddWin/raw_rangename");
		String name = fsaw_filename.getValue();

		Name namevar = book.createName();
		namevar.setNameName(name);
		namevar.setSheetIndex(sheetindex);
		CellRangeAddress cra = new CellRangeAddress(top, bottom, left, right);
		namevar.setRefersToFormula(cra.formatAsString());
		
		Window rangeAddWin = (Window) Path.getComponent("//p1/rangeAddWin");
		rangeAddWin.detach();
	}

	// dummy
	public void onRangeDelete() {
		try {
			Listbox rdw_rangeList = (Listbox) Path
					.getComponent("//p1/rangeDeleteWin/rdw_rangeList");

			Listitem st = rdw_rangeList.getSelectedItem();
			if (st != null) {
				String rangeName = st.getLabel();
				spreadsheet.getBook().removeName(rangeName);
			}

			Window rangeDeleteWin = (Window) Path
					.getComponent("//p1/rangeDeleteWin");
			rangeDeleteWin.detach();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	// dummy
	public void onRangeChoose() {
		try {
			Listbox rcw_rangeList = (Listbox) Path
					.getComponent("//p1/rangeChooseWin/rcw_rangeList");

			Listitem st = rcw_rangeList.getSelectedItem();
			Name range;
			Rect rect;
			if (st != null) {
				String rangeName = st.getLabel();
				range = spreadsheet.getBook().getName(rangeName);
				CellRangeAddress cra = CellRangeAddress.valueOf(range.getRefersToFormula());
				rect = new Rect(cra.getFirstColumn(), cra.getFirstRow(), 
						cra.getLastColumn(), cra.getLastRow());
				spreadsheet.setSelection(rect);
			}

			Window rangeChooseWin = (Window) Path
					.getComponent("//p1/rangeChooseWin");
			rangeChooseWin.detach();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void onRangeSelectChange(String type) {
		try {
			Listbox rangeList;

			if (type.equals("add")) {
				rangeList = (Listbox) Path
						.getComponent("//p1/rangeAddWin/raw_rangeList");
			} else if (type.equals("delete")) {
				rangeList = (Listbox) Path
						.getComponent("//p1/rangeDeleteWin/rdw_rangeList");
			} else if (type.equals("choose")) {
				rangeList = (Listbox) Path
						.getComponent("//p1/rangeChooseWin/rcw_rangeList");
			} else {
				return;
			}
			
			Listitem st = rangeList.getSelectedItem();
			Name range;
			Rect rect;
			if (st != null) {
				String rangeName = st.getLabel();
				range = spreadsheet.getBook().getName(rangeName);
				CellRangeAddress cra = CellRangeAddress.valueOf(range.getRefersToFormula());
				rect = new Rect(cra.getFirstColumn(), cra.getFirstRow(), 
						cra.getLastColumn(), cra.getLastRow());
				spreadsheet.setHighlight(rect);
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
