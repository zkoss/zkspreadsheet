/* DefaultHanderUtil.java
 * 
{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/5/15 , Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under Lesser GPL Version 2.1 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.ui.au.in;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.SortedSet;
import java.util.TreeSet;

import org.zkoss.json.JSONArray;
import org.zkoss.lang.Objects;
import org.zkoss.lang.Strings;
import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.util.Locales;
import org.zkoss.zss.api.AreaRef;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.impl.SheetImpl;
import org.zkoss.zss.model.SAutoFilter;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.SAutoFilter.FilterOp;
import org.zkoss.zss.model.SAutoFilter.NFilterColumn;
import org.zkoss.zss.model.SCell.CellType;
import org.zkoss.zss.model.sys.EngineFactory;
import org.zkoss.zss.model.sys.format.FormatContext;
import org.zkoss.zss.model.sys.format.FormatEngine;
import org.zkoss.zss.model.sys.format.FormatResult;
import org.zkoss.zss.range.SRange;
import org.zkoss.zss.range.SRanges;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * a util to help command to handle 'dirty' directly.
 * the dirty code are copy from Command implementation.
 * @author dennis
 *
 */
/*package*/ class AutoFilterDefaultHandler {

	private FilterRowInfo blankRowInfo;
	
	/*package*/ AreaRef processFilter(Spreadsheet spreadsheet,Sheet sheet,int row, int col, int field) {
		SSheet worksheet = ((SheetImpl)sheet).getNative();
		final SAutoFilter autoFilter = worksheet.getAutoFilter();
		final NFilterColumn filterColumn = autoFilter.getFilterColumn(field - 1,false);
		final String rangeAddr = autoFilter.getRegion().getReferenceString();
		final SRange range = SRanges.range(worksheet, rangeAddr);
		
		spreadsheet.smartUpdate("autoFilterPopup", 
			convertFilterInfoToJSON(row, col, field, rangeAddr, scanRows(field, filterColumn, range, worksheet)));
		
		AreaRef filterArea = new AreaRef(rangeAddr);
		return filterArea;
	}
	
	private Map convertFilterInfoToJSON(int row, int col, int field, String rangeAddr, SortedSet<FilterRowInfo> orderedRowInfos) {
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
	
	private SortedSet<FilterRowInfo> scanRows(int field, NFilterColumn fc, SRange range, SSheet worksheet) {
		SortedSet<FilterRowInfo> orderedRowInfos = new TreeSet<FilterRowInfo>(new FilterRowInfoComparator());
		
		blankRowInfo = new FilterRowInfo(BLANK_VALUE, "(Blanks)");
		final Set criteria1 = fc == null ? null : fc.getCriteria1();
		boolean hasBlank = false;
		boolean hasSelectedBlank = false;
		final int top = range.getRow() + 1;
		final int bottom = range.getLastRow();
		final int columnIndex = range.getColumn() + field - 1;
		FormatEngine fe = EngineFactory.getInstance().createFormatEngine();
		for (int i = top; i <= bottom; i++) {
			final SCell c = worksheet.getCell(i, columnIndex);
			if (!c.isNull() && c.getType() != CellType.BLANK) {
				FormatResult fr = fe.format(c, new FormatContext(ZssContext.getCurrent().getLocale()));
				String displaytxt = fr.getText();
				Object val = displaytxt;
				if(c.getType()==CellType.NUMBER && fr.isDateFormatted()){
					val = c.getDateValue();
				}
				
				FilterRowInfo rowInfo = new FilterRowInfo(val, displaytxt);
				//ZSS-299
				orderedRowInfos.add(rowInfo);
				if (criteria1 == null || criteria1.isEmpty() || criteria1.contains(displaytxt)) { //selected
					rowInfo.setSelected(true);
				}
			} else {
				hasBlank = true;
				boolean noFilterApplied = criteria1 == null || criteria1.isEmpty(); 
				if (!hasSelectedBlank && (noFilterApplied || criteria1.contains("="))) { //"=" means blank is selected
					blankRowInfo.setSelected(true);
					hasSelectedBlank = true;
				}
			}
		}
		if (hasBlank) {
			orderedRowInfos.add(blankRowInfo);
		}
		
		return orderedRowInfos;
	}
	
//	private static boolean isHiddenRow(int rowIdx, XSheet worksheet) {
//		final Row r = worksheet.getRow(rowIdx);
//		return r != null && r.getZeroHeight();
//	}
	
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

	/*package*/ void applyFilter(Spreadsheet spreadsheet, Sheet selectedSheet,
			String cellRangeAddr, boolean selectAll, int field, Object criteria) {
		final SRange range = SRanges.range(((SheetImpl)selectedSheet).getNative(), cellRangeAddr);
		
		if (selectAll) {
			range.enableAutoFilter(field, FilterOp.VALUES, null, null, true);
		} else { //partial selection
			JSONArray ary = (JSONArray) criteria;
			range.enableAutoFilter(field, FilterOp.VALUES, ary.toArray(new String[ary.size()]), null, true);
		}
	}
}
