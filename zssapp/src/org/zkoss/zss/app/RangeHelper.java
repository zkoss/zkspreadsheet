package org.zkoss.zss.app;

import org.zkoss.poi.ss.usermodel.Name;
//import org.zkoss.poi.ss.usermodel.Sheet;
//import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.zk.ui.Path;
//import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Listitem;
import org.zkoss.zul.Textbox;
import org.zkoss.zul.Window;

//help name range
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
//	public String[] getRangeList() {
//		Book book = MainWindowCtrl.getInstance().getSpreadsheet().getBook();
//		int len = book.getNumberOfNames();
//		String[] names = new String[len];
//		for(int j=0; j < len; ++j) {
//			names[j] = book.getNameAt(j).getNameName();
//		}
//		return names;
//	}

	// dummy
	public void onRangeAdd() {
		final Book book = spreadsheet.getBook();
		if (book == null) {
			return;
		}
		Sheet sheet = spreadsheet.getSelectedSheet();
		int sheetindex = book.getSheetIndex(sheet);
		Textbox fsaw_filename = (Textbox) Path
				.getComponent("//p1/rangeAddWin/raw_rangename");
		String name = fsaw_filename.getValue();

		Name namevar = book.getPoiBook().createName();
		namevar.setNameName(name);
		namevar.setSheetIndex(sheetindex);
		namevar.setRefersToFormula(Ranges.getAreaReference(spreadsheet.getSelection()));
		
		Window rangeAddWin = (Window) Path.getComponent("//p1/rangeAddWin");
		rangeAddWin.detach();
	}

	// dummy
	public void onRangeDelete() {
		final Book book = spreadsheet.getBook();
		if (book == null) {
			return;
		}
		try {
			Listbox rdw_rangeList = (Listbox) Path
					.getComponent("//p1/rangeDeleteWin/rdw_rangeList");

			Listitem st = rdw_rangeList.getSelectedItem();
			if (st != null) {
				String rangeName = st.getLabel();
				book.getPoiBook().removeName(rangeName);
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
		final Book book = spreadsheet.getBook();
		if (book == null) {
			return;
		}
		try {
			Listbox rcw_rangeList = (Listbox) Path
					.getComponent("//p1/rangeChooseWin/rcw_rangeList");

			Listitem st = rcw_rangeList.getSelectedItem();
			Name range;
			Rect rect;
			if (st != null) {
				String rangeName = st.getLabel();
				range = book.getPoiBook().getName(rangeName);
//				CellRangeAddress cra = CellRangeAddress.valueOf(range.getRefersToFormula());
				rect = new Rect(range.getRefersToFormula());
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
		final Book book = spreadsheet.getBook();
		if (book == null) {
			return;
		}
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
				range = book.getPoiBook().getName(rangeName);
//				CellRangeAddress cra = CellRangeAddress.valueOf(range.getRefersToFormula());
				rect = new Rect(range.getRefersToFormula());
				spreadsheet.setHighlight(rect);
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
