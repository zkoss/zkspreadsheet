/* CellMouseCommand.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Dec 18, 2007 12:10:40 PM     2007, Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under Lesser GPL Version 2.1 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.ui.au.in;


import static org.zkoss.zss.ui.au.in.Commands.parseKeys;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;

import org.zkoss.lang.Objects;
import org.zkoss.lang.Strings;
import org.zkoss.poi.ss.usermodel.AutoFilter;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.DateUtil;
import org.zkoss.poi.ss.usermodel.FilterColumn;
import org.zkoss.poi.ss.usermodel.RichTextString;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.zk.au.AuRequest;
import org.zkoss.zk.mesg.MZk;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zss.model.sys.XRange;
import org.zkoss.zss.model.sys.XRanges;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.model.sys.impl.BookHelper;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.CellMouseEvent;
import org.zkoss.zss.ui.event.FilterMouseEvent;
import org.zkoss.zss.ui.impl.XUtils;

/**
 * A Command (client to server) for handling user(client) start editing a cell
 * @author Dennis.Chen
 *
 */
public class CellMouseCommand implements Command {
	
	private FilterRowInfo blankRowInfo;
	
	//-- super --//
	public void process(AuRequest request) {
		final Component comp = request.getComponent();
		if (comp == null)
			throw new UiException(MZk.ILLEGAL_REQUEST_COMPONENT_REQUIRED, this);
		//final String[] data = request.getData();
		//TODO a little patch, "af" might shall be fired by a separate command?
		final Map data = (Map) request.getData();
		String type = (String) data.get("type");//command type
		if (data == null 
			|| (!"af".equals(type) && !"dv".equals(type) && data.size() != 9) 
			|| (("af".equals(type) || "dv".equals(type)) && data.size() != 10))
			throw new UiException(MZk.ILLEGAL_REQUEST_WRONG_DATA, new Object[] {Objects.toString(data), this});
		
		int shx = (Integer) data.get("shx");//x offset against spreadsheet
		int shy = (Integer) data.get("shy");
		int key = parseKeys((String) data.get("key"));
		String sheetId = (String) data.get("sheetId");
		int row = (Integer) data.get("row");
		int col = (Integer) data.get("col");
		int mx = (Integer) data.get("mx");//x offset against body
		int my = (Integer) data.get("my");
		
		Spreadsheet spreadsheet = (Spreadsheet) comp;
		XSheet sheet = ((Spreadsheet) comp).getSelectedXSheet();
		if (!XUtils.getSheetUuid(sheet).equals(sheetId))
			return;
		
		if ("lc".equals(type)) {
			type = org.zkoss.zss.ui.event.Events.ON_CELL_CLICK;
		} else if ("rc".equals(type)) {
			type = org.zkoss.zss.ui.event.Events.ON_CELL_RIGHT_CLICK;
		} else if ("dbc".equals(type)) {
			type = org.zkoss.zss.ui.event.Events.ON_CELL_DOUBLE_CLICK;
		} else if ("af".equals(type)) {
			type = org.zkoss.zss.ui.event.Events.ON_FILTER;
		} else if ("dv".equals(type)) {
			type = org.zkoss.zss.ui.event.Events.ON_VALIDATE_DROP;
		} else {
			throw new UiException("unknow type : " + type);
		}

		if (org.zkoss.zss.ui.event.Events.ON_FILTER.equals(type)) {
			int field = (Integer) data.get("field");
			processFilter(row, col, field, sheet, spreadsheet);
			Events.postEvent(new FilterMouseEvent(type, comp, shx, shy, key, sheet, row, col, mx, my, field));
		} else if (org.zkoss.zss.ui.event.Events.ON_VALIDATE_DROP.equals(type)) {
			int dvindex = (Integer) data.get("field");
			Events.postEvent(new FilterMouseEvent(type, comp, shx, shy, key, sheet, row, col, mx, my, dvindex));
		} else {
			Events.postEvent(new CellMouseEvent(type, comp, shx, shy, key, sheet, row, col, mx, my));
		}
	}
	
	private void processFilter (int row, int col, int field, XSheet worksheet, Spreadsheet spreadsheet) {
		final AutoFilter autoFilter = worksheet.getAutoFilter();
		final FilterColumn filterColumn = autoFilter.getFilterColumn(field - 1);
		final String rangeAddr = autoFilter.getRangeAddress().formatAsString();
		final XRange range = XRanges.range(worksheet, rangeAddr);
		
		spreadsheet.smartUpdate("autoFilterPopup", 
			convertFilterInfoToJSON(row, col, field, rangeAddr, scanRows(field, filterColumn, range, worksheet)));
	}
	
	private Map convertFilterInfoToJSON(int row, int col, int field, String rangeAddr, TreeSet<FilterRowInfo> orderedRowInfos) {
		final Map data = new HashMap();
		
		boolean selectAll = true;
		boolean select = false;
		final ArrayList<Map> sortedItems = new ArrayList<Map>();
		for (FilterRowInfo info : orderedRowInfos) {			
			if (info == blankRowInfo) {
				data.put("blank", info.seld);
				if (info.seld) {
					select = true;
				} else {
					selectAll = false;
				}
			} else {
				HashMap item = new HashMap();
				sortedItems.add(item);
				item.put("v", info.display);
				if (info.isSelected()) {
					item.put("s", "t"); //selected, "t" stand for true
					select = true;
				} else {
					selectAll = false;
				}
			}
		}
		data.put("items", sortedItems);
		data.put("row", row);
		data.put("col", col);
		data.put("field", field);
		data.put("range", rangeAddr);
		data.put("select", selectAll ? "all" : select ? "mix" : "none"); 
		return data;
	}
	
	private TreeSet<FilterRowInfo> scanRows(int field, FilterColumn fc, XRange range, XSheet worksheet) {
		TreeSet<FilterRowInfo> orderedRowInfos = new TreeSet<FilterRowInfo>(new FilterRowInfoComparator());
		
		blankRowInfo = new FilterRowInfo(BLANK_VALUE, "(Blanks)");
		final Set criteria1 = fc == null ? null : fc.getCriteria1();
		final boolean nofilter = criteria1 == null || criteria1.isEmpty(); 
		boolean hasBlank = false;
		boolean selectedBlank = false;
		final int top = range.getRow() + 1;
		final int bottom = range.getLastRow();
		final int columnIndex = range.getColumn() + field - 1;
		for (int i = top; i <= bottom; i++) {
			if (nofilter && isHiddenRow(i, worksheet)) {
				continue;
			}
			final Cell c = XUtils.getCell(worksheet, i, columnIndex);
			final boolean blankcell = BookHelper.isBlankCell(c);
			if (!blankcell) {
				String displaytxt = BookHelper.getCellText(c);
				Object val = BookHelper.getEvalCellValue(c);
				if (val instanceof RichTextString) {
					val = ((RichTextString)val).getString();
				} else if (c.getCellType() == Cell.CELL_TYPE_NUMERIC && DateUtil.isCellDateFormatted(c)) {
					val = c.getDateCellValue();
				}
				FilterRowInfo rowInfo = new FilterRowInfo(val, displaytxt); 
				orderedRowInfos.add(rowInfo);
				if (criteria1 == null || criteria1.isEmpty() || criteria1.contains(displaytxt)) { //selected
					rowInfo.setSelected(true);
				}
			} else {
				hasBlank = true;
				if (!selectedBlank && (nofilter || criteria1.contains("="))) { //selected
					blankRowInfo.setSelected(true);
				}
			}
		}
		if (hasBlank) {
			orderedRowInfos.add(blankRowInfo);
		}
		
		return orderedRowInfos;
	}
	
	private static boolean isHiddenRow(int rowIdx, XSheet worksheet) {
		final Row r = worksheet.getRow(rowIdx);
		return r != null && r.getZeroHeight();
	}
	
	private final static Comparable BLANK_VALUE = new Comparable() {
		@Override
		public int compareTo(Object o) {
			return BLANK_VALUE.equals(o) ? 0 : 1; //unless same otherwise BLANK_VALUE is always the biggest!
		}
	};
	
	private static class FilterRowInfo {
		private Object value;
		private String display;
		private boolean seld;
		
		FilterRowInfo(Object val, String displayVal) {
			value = val;
			display = displayVal;
		}
		
		Object getValue() {
			return value;
		}
		
		String getDisplay() {
			return display;
		}
		
		void setSelected(boolean selected) {
			seld = selected;
		}
		
		boolean isSelected() {
			return seld;
		}

		public int hashCode() {
			return value == null ? 0 : value.hashCode();
		}

		@Override
		public boolean equals(Object obj) {
			if (this == obj)
				return true;
			if (!(obj instanceof FilterRowInfo))
				return false;
			final FilterRowInfo other = (FilterRowInfo) obj;
			return Objects.equals(other.value, this.value);
		}
	}
	
	private static class FilterRowInfoComparator implements Comparator<FilterRowInfo> {		
		@Override
		public int compare(FilterRowInfo o1, FilterRowInfo o2) {
			final Object val1 = o1.value;
			final Object val2 = o2.value;
			final int type1 = getType(val1);
			final int type2 = getType(val2);
			final int typediff = type1 - type2;
			if (typediff != 0) {
				return typediff;
			}
			switch(type1) {
			case 1: //Date
				return compareDates((Date)val1, (Date)val2);
			case 2: //Number
				return ((Double)val1).compareTo((Double)val2);
			case 3: //String
				return ((String)val1).compareTo((String)val2);
			case 4: //Boolean
				final boolean b1 = ((Boolean)val1).booleanValue();
				final boolean b2 = ((Boolean)val2).booleanValue();
				return !b1 && b2 ? -1 : b1 && !b2 ? 1 : 0;
			case 5: //Error(Byte)
				return ((Byte)val1).compareTo((Byte)val2);
			default:
			case 6: //(Blanks)
				return 0;
			}
		}
		private int compareDates(Date val1, Date val2) {
			final Calendar cal1 = Calendar.getInstance();
			final Calendar cal2 = Calendar.getInstance();
			cal1.setTime((Date)val1);
			cal2.setTime((Date)val2);
			
			//year
			final int y1 = cal1.get(Calendar.YEAR);
			final int y2 = cal2.get(Calendar.YEAR);
			final int ydiff = y2 - y1; //bigger year is less in sorting
			if (ydiff != 0) {
				return ydiff;
			}
			
			//month
			final int m1 = cal1.get(Calendar.MONTH);
			final int m2 = cal2.get(Calendar.MONTH);
			final int mdiff = m1 - m2; 
			if (mdiff != 0) {
				return mdiff;
			}
			
			//day
			final int d1 = cal1.get(Calendar.DAY_OF_MONTH);
			final int d2 = cal2.get(Calendar.DAY_OF_MONTH);
			final int ddiff = d1 - d2; //smaller month is bigger in sorting 
			if (ddiff != 0) {
				return ddiff;
			}
			
			//hour
			final int h1 = cal1.get(Calendar.HOUR_OF_DAY);
			final int h2 = cal2.get(Calendar.HOUR_OF_DAY);
			final int hdiff = h1 - h2;
			if (hdiff != 0) {
				return hdiff;
			}
			
			//minutes
			final int mm1 = cal1.get(Calendar.MINUTE);
			final int mm2 = cal2.get(Calendar.MINUTE);
			final int mmdiff = mm1 - mm2;
			if (mmdiff != 0) {
				return mmdiff;
			}
			
			//seconds
			final int s1 = cal1.get(Calendar.SECOND);
			final int s2 = cal2.get(Calendar.SECOND);
			final int sdiff = s1 - s2;
			if (sdiff != 0) {
				return sdiff;
			}
			
			//millseconds
			final int ms1 = cal1.get(Calendar.MILLISECOND);
			final int ms2 = cal2.get(Calendar.MILLISECOND);
			return ms1 - ms2;
		}
		//Date < Number < String < Boolean(FALSE < TRUE) < Error(byte) < (Blanks)
		private int getType(Object val) {
			if (val instanceof Date) {
				return 1;
			}
			if (val instanceof Byte) { //error
				return 5;
			}
			if (val instanceof Number) {
				return 2;
			}
			if (val instanceof String) {
				return Strings.isEmpty((String)val) ? 6 : 3;
			}
			if (val instanceof Boolean) {
				return 4;
			}
			return 6;
		}
	}
}