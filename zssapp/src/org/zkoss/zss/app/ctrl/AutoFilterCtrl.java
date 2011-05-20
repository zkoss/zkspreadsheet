/* AutoFilterCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		May 13, 2011 10:24:13 AM , Created by sam
}}IS_NOTE

Copyright (C) 2011 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.app.ctrl;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Set;
import java.util.TreeSet;

import org.zkoss.lang.Objects;
import org.zkoss.poi.ss.usermodel.AutoFilter;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.zul.Dialog;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zul.Button;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Listitem;
import org.zkoss.zul.ListitemRenderer;

/**
 * @author sam
 *
 */
public class AutoFilterCtrl extends GenericForwardComposer {
	
	private Dialog _autoFilterDialog;
	private Listbox _filterListbox;
	
	private Button okBtn;
	
	private final static String KEY_ROW_INFO = "rowInfo";
	private final static String BLANK_DISPLAY = "Blanks";
	private final static Comparable BLANK_VALUE = new Comparable() {
		@Override
		public int compareTo(Object o) {
			return BLANK_VALUE.equals(o) ? 0 : 1; //unless same otherwise BLANK_VALUE is always the biggest!
		}
	}; 
	private final RowInfo BLANK_ROW_INFO = new RowInfo(0, 0, BLANK_VALUE, BLANK_DISPLAY);

	int fieldIndex = 0;
	int columnIndex = 0;
	private Worksheet worksheet;
	private Range range;
	
	private boolean isHiddenRow(Worksheet sheet, int rowIdx) {
		final Row r = sheet.getRow(rowIdx);
		return r != null && r.getZeroHeight();
	}
	
	public void onOpen$_autoFilterDialog(ForwardEvent evt) {
		Object[] info = (Object[]) evt.getOrigin().getData();
		fieldIndex = (Integer) info[0];
		columnIndex = (Integer) info[1];
		range = (Range) info[2];

		boolean hasBlank = false;
		TreeSet<RowInfo> rowInfos = new TreeSet<RowInfo>();

		worksheet = (Worksheet) range.getSheet();
		final int top = range.getRow() + 1;
		final int bottom = range.getLastRow();
		for (int i = top; i <= bottom; i++) {
			final Cell c = Utils.getCell(worksheet, i, columnIndex);
			final boolean blankcell = BookHelper.isBlankCell(c);
			if (!blankcell) {
				String displaytxt = BookHelper.getCellText(c);
				Object val = BookHelper.getCellValue(c); 
				rowInfos.add(new RowInfo(i, columnIndex, val, displaytxt));
			} else {
				hasBlank = true;
			}
		}
		if (hasBlank)
			rowInfos.add(BLANK_ROW_INFO);

		_filterListbox.setModel(new ListModelList(rowInfos));
		
		//handle selection
		ListModelList model = (ListModelList) _filterListbox.getListModel();
		for(final Iterator it = model.iterator(); it.hasNext();) {
			final RowInfo rowInfo = (RowInfo) it.next();
			if (!isHiddenRow(worksheet, rowInfo.row))
				model.addSelection(rowInfo);
		}
		
		_filterListbox.setItemRenderer(new ListitemRenderer() {
			@Override
			public void render(Listitem item, Object data) throws Exception {
				RowInfo info = (RowInfo)data;
				item.setLabel(info.display);
			}
		});
	}
	
	private static class RowInfo implements Comparable {
		int row;
		int col;
		Object value;
		String display;
		
		RowInfo(int rowIdx, int colIdx, Object val, String displayVal) {
			row = rowIdx;
			col = colIdx;
			value = val;
			display = displayVal;
		}

		public int hashCode() {
			return value == null ? 0 : value.hashCode();
		}

		@Override
		public boolean equals(Object obj) {
			if (this == obj)
				return true;
			if (!(obj instanceof RowInfo))
				return false;
			final RowInfo other = (RowInfo) obj;
			return Objects.equals(other.value, this.value);
		}

		@Override
		public int compareTo(Object o) {
			final RowInfo other = (RowInfo) o;
			return ((Comparable)this.value).compareTo(other.value);
		}
	}

	public void onClick$okBtn() {
		final ListModelList model = (ListModelList) _filterListbox.getListModel();
		final int itemcount = model.size();
		final Set selected = model.getSelection();
		final int selcount = selected.size();
		if (selcount < itemcount) { //partial selection
			final String[] criteria = new String[selcount];
			int j = 0;
			for (final Iterator it = new TreeSet(selected).iterator(); it.hasNext(); ) {
				RowInfo info = (RowInfo) it.next();
				criteria[j++] = BLANK_ROW_INFO.equals(info) ? "=" : info.display;
			}
			range.autoFilter(fieldIndex, criteria, AutoFilter.FILTEROP_VALUES, null, null);
		} else { //select all!
			range.autoFilter(fieldIndex, null, AutoFilter.FILTEROP_VALUES, null, null);
		}
		_autoFilterDialog.fireOnClose(null);
	}
}