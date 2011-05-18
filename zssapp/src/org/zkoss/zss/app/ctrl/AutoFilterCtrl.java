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
import java.util.List;
import java.util.TreeSet;

import org.zkoss.poi.ss.usermodel.AutoFilter;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.zul.Dialog;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.Worksheet;
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
	private final RowInfo BLANK_ROW_INFO = new RowInfo(0, 0, "=", BLANK_DISPLAY);

	int fieldIndex = 0;
	int columnIndex = 0;
	private Worksheet worksheet;
	private Range range;
	
	public void onOpen$_autoFilterDialog(ForwardEvent evt) {
		Object[] info = (Object[]) evt.getOrigin().getData();
		fieldIndex = (Integer) info[0];
		columnIndex = (Integer) info[1];
		range = (Range) info[2];

		boolean hasBlank = false;
		TreeSet<RowInfo> rowInfos = new TreeSet<RowInfo>();

		worksheet = (Worksheet) range.getSheet();
		for (int i = range.getRow() + 1; i <= range.getLastRow(); i++) {
			Cell c = Utils.getCell(worksheet, i, columnIndex);
			if (c != null) {
				int type = c.getCellType();
				if (type == Cell.CELL_TYPE_BLANK) {
					hasBlank = true;
				} else {
					String val = Ranges.range(worksheet, i, columnIndex).getFormatText().getCellFormatResult().text;
					rowInfos.add(new RowInfo(i, columnIndex, val, val));
				}
			} else {
				hasBlank = true;
			}
		}
		if (hasBlank)
			rowInfos.add(BLANK_ROW_INFO);

		_filterListbox.setModel(new ListModelList(rowInfos));
		_filterListbox.setItemRenderer(new ListitemRenderer() {
			@Override
			public void render(Listitem item, Object data) throws Exception {
				RowInfo info = (RowInfo)data;
				if (info.display != null)
					item.setLabel(info.display);
				item.setAttribute(KEY_ROW_INFO, info);
				
				if (info == BLANK_ROW_INFO) {
					setBlankRowSelection(item);
				} else {
					Row r = worksheet.getRow(info.row);
					if (r != null) {
						item.setSelected(!r.getZeroHeight());
					}	
				}
			}
			private boolean isRowHidden(Worksheet sheet, int rowIdx, int colIdx) {
				Row r = sheet.getRow(rowIdx);
				if (r == null)
					r = Utils.getOrCreateRow(sheet, rowIdx);
				return r.getZeroHeight();
			}
			private void setBlankRowSelection(Listitem item) {
				boolean isHidden = true;
				for (int i = range.getRow(); i <= range.getLastRow(); i++) {
					Cell c = Utils.getCell(worksheet, i, columnIndex);
					if (c != null && c.getCellType() != Cell.CELL_TYPE_BLANK)
						continue;
					if ( !(isHidden = isRowHidden(worksheet, i, columnIndex)) ) {
						break;
					}
				}
				item.setSelected(!isHidden);
			}
		});
	}
	
	class RowInfo implements Comparable {
		int row;
		int col;
		String text;
		String display;
		
		RowInfo(int rowIdx, int colIdx, String val, String displayVal) {
			row = rowIdx;
			col = colIdx;
			text = val;
			display = displayVal;
		}

		public int hashCode() {
			final int prime = 31;
			int result = 1;
			result = prime * result + getOuterType().hashCode();
			result = prime * result
					+ ((display == null) ? 0 : display.hashCode());
			return result;
		}

		@Override
		public boolean equals(Object obj) {
			if (this == obj)
				return true;
			if (obj == null)
				return false;
			if (getClass() != obj.getClass())
				return false;
			RowInfo other = (RowInfo) obj;
			if (!getOuterType().equals(other.getOuterType()))
				return false;
			if (display == null) {
				if (other.display != null)
					return false;
			} else if (!display.equals(other.display))
				return false;
			return true;
		}

		private AutoFilterCtrl getOuterType() {
			return AutoFilterCtrl.this;
		}

		@Override
		public int compareTo(Object o) {
			if (display == BLANK_DISPLAY)
				return Integer.MAX_VALUE;
			RowInfo target = (RowInfo) o;
			return display.compareTo(target.display);
		}
	}

	public void onClick$okBtn() {
		List items = _filterListbox.getItems();
		ArrayList<String> criteria = new ArrayList<String>();
		for (int i = 0; i < items.size(); i++) {
			Listitem item = (Listitem) items.get(i);
			RowInfo info = (RowInfo) item.getAttribute(KEY_ROW_INFO);
			if (item.isSelected())
				criteria.add(info.text);
		}
		range.autoFilter(fieldIndex, criteria.toArray(new String[criteria.size()]), AutoFilter.FILTEROP_VALUES, null, true);
		_autoFilterDialog.fireOnClose(null);
	}
}