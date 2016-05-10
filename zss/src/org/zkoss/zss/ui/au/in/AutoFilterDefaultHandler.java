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

import java.io.Serializable;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.SortedSet;
import java.util.TreeSet;
import java.util.LinkedHashSet;

import org.zkoss.json.JSONArray;
import org.zkoss.lang.Objects;
import org.zkoss.lang.Strings;
import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.util.Locales;
import org.zkoss.zss.api.AreaRef;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.impl.SheetImpl;
import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.ErrorValue;
import org.zkoss.zss.model.SAutoFilter;
import org.zkoss.zss.model.SAutoFilter.SColorFilter;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SCellStyle;
import org.zkoss.zss.model.SExtraStyle;
import org.zkoss.zss.model.SFill.FillPattern;
import org.zkoss.zss.model.SFont;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.SAutoFilter.FilterOp;
import org.zkoss.zss.model.SAutoFilter.NFilterColumn;
import org.zkoss.zss.model.SCell.CellType;
import org.zkoss.zss.model.STable;
import org.zkoss.zss.model.SFill;
import org.zkoss.zss.model.sys.EngineFactory;
import org.zkoss.zss.model.sys.format.FormatContext;
import org.zkoss.zss.model.sys.format.FormatEngine;
import org.zkoss.zss.model.sys.format.FormatResult;
import org.zkoss.zss.model.impl.AbstractAutoFilterAdv;
import org.zkoss.zss.model.impl.AbstractSheetAdv;
import org.zkoss.zss.model.impl.FillImpl;
import org.zkoss.zss.model.impl.FontImpl;
import org.zkoss.zss.range.SRange;
import org.zkoss.zss.range.SRanges;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * a util to help command to handle 'dirty' directly.
 * the dirty code are copy from Command implementation.
 * @author dennis
 *
 */
/*package*/ class AutoFilterDefaultHandler implements Serializable {
	private static final long serialVersionUID = -8371978786540085707L;
	private static final SFill BLANK_FILL = FillImpl.BLANK_FILL; //ZSS-1191
	private static final SFont BLANK_FONT = FontImpl.BLANK_FONT; //ZSS-1191
	
	private FilterRowInfo blankRowInfo;
	
	/*package*/ AreaRef processFilter(Spreadsheet spreadsheet,Sheet sheet,int row, int col, int field) {
		SSheet worksheet = ((SheetImpl)sheet).getNative();
		//ZSS-988
		STable table = ((AbstractSheetAdv)worksheet).getTableByRowCol(row, col);
		final SAutoFilter autoFilter = table == null ? worksheet.getAutoFilter() : table.getAutoFilter();
		
		final NFilterColumn filterColumn = autoFilter.getFilterColumn(field - 1,false);
		String rangeAddr = autoFilter.getRegion().getReferenceString();
		final SRange range = SRanges.range(worksheet, rangeAddr);

		//ZSS-1191
		final SColorFilter colorFilter = filterColumn == null ? null : filterColumn.getColorFilter();
		final SExtraStyle extraStyle = colorFilter == null ? null : colorFilter.getExtraStyle();
		final SFill filterFill = extraStyle == null ? null : extraStyle.getFill();
		final boolean byFontColor = colorFilter == null ? false : colorFilter.isByFontColor();
		
		//ZSS-704: Note that scanRows() could provide new bottom
		final int bottom = range.getLastRow();
		Object[] results = scanRows(field, filterColumn, range, worksheet, 
				table, filterFill, byFontColor); //ZSS-988, ZSS-1191
		@SuppressWarnings("unchecked")
		SortedSet<FilterRowInfo> orderedRowInfos = (SortedSet<FilterRowInfo>) results[0];
		if (bottom != ((Integer) results[1]).intValue()) {
			CellRegion region = new CellRegion(range.getRow(), range.getColumn(), bottom, range.getLastColumn()); 
			rangeAddr = region.getReferenceString(); 
		}
		
		//ZSS-1191
		int type = ((Integer)results[3]).intValue();
		Set<SFill> ccitems = (Set<SFill>) results[2];
		Set<SFill> fcitems = (Set<SFill>) results[4];
		
		spreadsheet.smartUpdate("autoFilterPopup", 
			convertFilterInfoToJSON(row, col, field, rangeAddr, orderedRowInfos,
					type, ccitems, filterFill, fcitems, byFontColor)); //ZSS-1191
		
		AreaRef filterArea = new AreaRef(rangeAddr);
		
		return filterArea;
	}
	
	//see Popup.js#zssex.AutoFilterPopup 
	private Map convertFilterInfoToJSON(int row, int col, int field,
			String rangeAddr, SortedSet<FilterRowInfo> orderedRowInfos,
			int type, Set<SFill> ccitems, SFill filterFill, Set<SFill> fcitems, 
			boolean byFontColor) { //ZSS-1191
		final Map data = new HashMap();
		
		boolean selectAll = filterFill == null; //ZSS-1191
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
		
		//ZSS-1191
		if (filterFill != null) {
			String fg = filterFill.getFillColor().getHtmlColor();
			String bg = filterFill.getBackColor().getHtmlColor();
			
			//ZSS-1911: in filter by CELL_COLOR; fg/bg should reversed to follow SCellStyle
			if (filterFill.getFillPattern() == FillPattern.SOLID && !byFontColor) {
				final String tmp = fg;
				fg = bg;
				bg = tmp;
			}
			filterFill = new FillImpl(filterFill.getFillPattern(), fg, bg);
			data.put("colorpat", filterFill.getFillPattern().name());
			data.put("colorfg", fg);
			data.put("colorbg", bg);
			data.put("fontColor", byFontColor);
		}
		
		final ArrayList<Map> ccfills = new ArrayList<Map>();
		for (SFill fill : ccitems) {
			HashMap item = new HashMap();
			ccfills.add(item);
			if (fill.getFillPattern() == FillPattern.NONE) {
				item.put("t", "No Fill");
			} else {
				item.put("t", "");
				if (fill.getFillPattern() == FillPattern.SOLID) {
					item.put("bk", "background: " + fill.getBackColor().getHtmlColor());
				} else {
					item.put("bk", ((FillImpl)fill).getFillPatternHtml());
				}
			}
			if (fill.equals(filterFill)) {
				item.put("s", true);
			}
			item.put("pat", fill.getFillPattern().name());
			item.put("fg", fill.getFillColor().getHtmlColor());
			item.put("bg", fill.getBackColor().getHtmlColor());
		}
		
		final ArrayList<Map> fcfills = new ArrayList<Map>();
		for (SFill fill : fcitems) {
			HashMap item = new HashMap();
			fcfills.add(item);
			if (fill.getFillPattern() == FillPattern.NONE) {
				item.put("t", "Auto");
			} else {
				item.put("t", "");
				item.put("bk", "background: " + fill.getFillColor().getHtmlColor());
			}
			if (fill.equals(filterFill)) {
				item.put("s", true);
			}
			item.put("pat", fill.getFillPattern().name());
			item.put("fg", fill.getFillColor().getHtmlColor());
			item.put("bg", fill.getBackColor().getHtmlColor());
		}

		data.put("ccitems", ccfills);
		data.put("fcitems", fcfills);
		data.put("type", type);
		
		data.put("items", sortedItems);
		data.put("row", row);
		data.put("col", col);
		data.put("field", field);
		data.put("range", rangeAddr);
		data.put("select", selectAll ? "all" : select ? "mix" : "none"); 
		return data;
	}

	// ZSS-704
	// [0]: SortedSet<FilterRowInfo>; 
	// [1]: new bottom; 
	// [2]: LinkedHashSet<SFill> for CELL_COLOR; 
	// [3]: which kind of filter(Number Filter/ Date Filter/ Text Filter);
	// [4]: LinkedHashSet<SFill> for FONT_COLOR 
	private Object[] scanRows(int field, NFilterColumn fc, SRange range, 
			SSheet worksheet, STable table, SFill filterFill, boolean byFontColor) { //ZSS-988
		SortedSet<FilterRowInfo> orderedRowInfos = new TreeSet<FilterRowInfo>(new FilterRowInfoComparator());
		
		//ZSS-1191
		LinkedHashSet<SFill> ccitems = new LinkedHashSet<SFill>(); //ZSS-1191: CELL_COLOR
		LinkedHashSet<SFill> fcitems = new LinkedHashSet<SFill>(); //ZSS-1191: FONT_COLOR
		int[] types = new int[] {0, 0, 0}; //0: date, 1: number, 2: string
		int candidate = 0; // which kind of filter (DateFilter/NumberFilter/TextFilter)
		
		blankRowInfo = new FilterRowInfo(BLANK_VALUE, "(Blanks)");
		final Set criteria1 = fc == null ? null : fc.getCriteria1();
		boolean hasBlank = false;
		boolean hasSelectedBlank = false;
		final int top = range.getRow() + 1;
		int bottom = range.getLastRow();
		final int columnIndex = range.getColumn() + field - 1;
		final SFont defaultFont = worksheet.getBook().getDefaultFont(); //ZSS-1191
		FormatEngine fe = EngineFactory.getInstance().createFormatEngine();
		for (int i = top; i <= bottom; i++) {
			//ZSS-988: filter column with no criteria should not show option of hidden row 
			if ((criteria1 == null || criteria1.isEmpty()) && filterFill == null) {
				if (worksheet.getRow(i).isHidden())
					continue;
			}
			final SCell c = worksheet.getCell(i, columnIndex);
			
			//ZSS-1191
			final SCellStyle style = c.isNull() ? null : c.getCellStyle();
			final SFill ccfill = style == null ? BLANK_FILL : style.getFill();
			ccitems.add(ccfill);
			final SFont font = style == null ? null : style.getFont();
			final boolean isDefaultFont = defaultFont.equals(font);
			final SFill fcfill = font == null || isDefaultFont ? 
				BLANK_FILL : new FillImpl(FillPattern.SOLID, font.getColor(), null);
			fcitems.add(fcfill);			
			int type0 = 2;
			
			if (!c.isNull() && c.getType() != CellType.BLANK) {
				FormatResult fr = fe.format(c, new FormatContext(ZssContext.getCurrent().getLocale()));
				String displaytxt = fr.getText();
				if(!hasBlank && displaytxt.trim().isEmpty()) { //ZSS-707: show as blank; then it is blank
					hasBlank = true;
					hasSelectedBlank = prepareBlankRow(criteria1, hasSelectedBlank, filterFill); //ZSS-1191
				} else {				
					Object val = c.getValue(); // ZSS-707
					if(c.getType()==CellType.NUMBER && fr.isDateFormatted()){
						val = c.getDateValue();
					}
					
					//ZSS-1191: Date 1, Number 2, String 3, Boolean 4 is Number; Error 5 and Blank 6 is String.
					final int type = getType(val);
					type0 = (type == 4 ? 2 : type >= 5 ? 3 : type) - 1;
					
					FilterRowInfo rowInfo = new FilterRowInfo(val, displaytxt);
					//ZSS-299
					orderedRowInfos.add(rowInfo);
					if (filterFill == null) { //ZSS-1191: color filter excludes value filter
						if (criteria1 == null || criteria1.isEmpty() || criteria1.contains(displaytxt)) { //selected
							rowInfo.setSelected(true);
						}
					}
				}
			} else if (!hasBlank){
				hasBlank = true;
				hasSelectedBlank = prepareBlankRow(criteria1, hasSelectedBlank, filterFill); //ZSS-1191
			}
			
			//ZSS-1191: Date 0, Number 1, String 2
			types[type0] = types[type0] + 1;
			if (types[type0] > types[candidate]) {
				candidate = type0;
			}
		}
		//ZSS-988: Only when it is not a table filter, it is possible to change the last row.
		if (table == null) {
			//ZSS-988: when hit Table cell; must stop
			int blm = Integer.MAX_VALUE;
			final SSheet sheet = range.getSheet();
			for (STable tb : sheet.getTables()) {
				final CellRegion rgn = tb.getAllRegion().getRegion();
				final int l = rgn.getColumn();
				final int r = rgn.getLastColumn();
				final int t = rgn.getRow();
				if (l <= columnIndex && columnIndex <= r && t > bottom && blm >= t)
					blm = t - 1;
			}

			final int maxblm = Math.min(blm, worksheet.getEndRowIndex());
			//ZSS-704: user could have enter non-blank value along the filter, must add that into
			final int left = range.getColumn();
			final int right = range.getLastColumn();
			boolean leaveLoop = false;
			for (int i = bottom+1; i <= maxblm ; ++i) {
				final SCell c = worksheet.getCell(i, columnIndex);
				
				//ZSS-1191
				final SCellStyle style = c.isNull() ? null : c.getCellStyle();
				final SFill ccfill = style == null ? BLANK_FILL : style.getFill();
				ccitems.add(ccfill);
				final SFont font = style == null ? null : style.getFont();
				final SFill fcfill = font == null ? BLANK_FILL : new FillImpl(FillPattern.SOLID, font.getColor(), null);
				fcitems.add(fcfill);
				int type0 = 2;

				if (!c.isNull() && c.getType() != CellType.BLANK) {
					FormatResult fr = fe.format(c, new FormatContext(ZssContext.getCurrent().getLocale()));
					String displaytxt = fr.getText();
					if(!hasBlank && displaytxt.trim().isEmpty()) { //ZSS-707: show as blank; then it is blank
						hasBlank = true;
						hasSelectedBlank = prepareBlankRow(criteria1, hasSelectedBlank, filterFill); //ZSS-1191
					} else {
						Object val = c.getValue(); // ZSS-707
						if(c.getType()==CellType.NUMBER && fr.isDateFormatted()){
							val = c.getDateValue();
						}
						
						//ZSS-1191: Date 1, Number 2, String 3, Boolean 4 is Number; Error 5 and Blank 6 is String.
						final int type = getType(val);
						type0 = (type == 4 ? 2 : type >= 5 ? 3 : type) - 1;
						
						FilterRowInfo rowInfo = new FilterRowInfo(val, displaytxt);
						//ZSS-299
						orderedRowInfos.add(rowInfo);
						if (filterFill == null) { //ZSS-1191: color filter excludes value filter
							if (criteria1 == null || criteria1.isEmpty() || criteria1.contains(displaytxt)) { //selected
								rowInfo.setSelected(true);
							}
						}
					}
				} else {
					//really an empty cell?
					int[] ltrb = getMergedMinMax(worksheet, i, columnIndex);
					if (ltrb == null) {
						if (neighborIsBlank(worksheet, left, right, i, columnIndex)) {
							bottom = i - 1;
							leaveLoop = true;
						}
					} else {
						i = ltrb[3];
					}
					if (!hasBlank) {
						hasBlank = true;
						hasSelectedBlank = prepareBlankRow(criteria1, hasSelectedBlank, filterFill); //ZSS-1191
					}
				}
				
				//ZSS-1191: Date 0, Number 1, String 2
				types[type0] = types[type0] + 1;
				if (types[type0] > types[candidate]) {
					candidate = type0;
				}
				
				if (leaveLoop) {
					break;
				}
			}
		}
		if (hasBlank) {
			orderedRowInfos.add(blankRowInfo);
		}
		return new Object[] {orderedRowInfos, bottom, 
				ccitems.size() > 1 ? ccitems : Collections.EMPTY_SET, 
				new Integer(candidate+1),
				fcitems.size() > 1 ? fcitems : Collections.EMPTY_SET};
	}

	//ZSS-707
	private boolean prepareBlankRow(Set criteria1, boolean hasSelectedBlank, SFill filterFill) {
		boolean noFilterApplied = criteria1 == null || criteria1.isEmpty(); 
		if (filterFill == null) { //ZSS-1911: color filter excludes value filter
			if (!hasSelectedBlank && (noFilterApplied || criteria1.contains("="))) { //"=" means blank is selected
				blankRowInfo.setSelected(true);
				return true;
			}
		}
		return hasSelectedBlank;
	}
	//ZSS-704
	// return null if not merged cell or the merged cell is blank; return
	// merged l, t, r, b if exists
	private int[] getMergedMinMax(SSheet worksheet, int row, int col) {
		CellRegion merged = worksheet.getMergedRegion(row, col);
		if (merged == null) {
			return null;
		} else {
			int l = merged.getColumn();
			int t = merged.getRow();
			final SCell c0 = worksheet.getCell(t, l);
			if (!c0.isNull() && c0.getType() != CellType.BLANK) { // non empty merged cell
				return new int[] {l, t, merged.getLastColumn(), merged.getLastRow()};
			} else {
				return null;
			}
		}
	}
	
	//ZSS-704
	// whether neighbor cell between left and right is blank.
	private boolean neighborIsBlank(SSheet sheet, int left, int right, int row, int col) {
		for (int j = left; j <= right; ++j) {
			if (j == col) continue;
			final SCell c = sheet.getCell(row, j);
			if (!c.isNull() && c.getType() != CellType.BLANK) {
				return false;
			} else {
				CellRegion merged = sheet.getMergedRegion(row, j);
				if (merged != null) {
					int l = merged.getColumn();
					int t = merged.getRow();
					final SCell c0 = sheet.getCell(t, l);
					if (!c0.isNull() && c0.getType() != CellType.BLANK) {
						return false;
					}
				}
			}
		}
		return true;
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
				//ZSS-935
				final byte by1 = val1 instanceof ErrorValue ? Byte.valueOf(((ErrorValue)val1).getCode()) : ((Byte) val1).byteValue();
				final byte by2 = val2 instanceof ErrorValue ? Byte.valueOf(((ErrorValue)val2).getCode()) : ((Byte) val2).byteValue();
				return by1 - by2;
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
	}
	
	//ZSS-1191
	//Date < Number < String < Boolean(FALSE < TRUE) < Error(byte) < (Blanks)
	private static int getType(Object val) {
		if (val instanceof Date) {
			return 1;
		}
		if (val instanceof ErrorValue || val instanceof Byte) { //error, ZSS-707
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
 
	/*package*/ void applyFilter(Spreadsheet spreadsheet, Sheet selectedSheet,
			String cellRangeAddr, boolean selectAll, int field, Object criteria, FilterOp op) { //ZSS-1191
		final SRange range = SRanges.range(((SheetImpl)selectedSheet).getNative(), cellRangeAddr);
		
		if (FilterOp.CELL_COLOR == op) { //ZSS-1191
			JSONArray ary = (JSONArray) criteria;
			range.enableAutoFilter(field, FilterOp.CELL_COLOR, ary.toArray(new String[ary.size()]), null, true);
		} else if (FilterOp.FONT_COLOR == op) { //ZSS-1191
			JSONArray ary = (JSONArray) criteria;
			range.enableAutoFilter(field, FilterOp.FONT_COLOR, ary.toArray(new String[ary.size()]), null, true);
		} else {
			if (selectAll) {
				range.enableAutoFilter(field, FilterOp.VALUES, null, null, true);
			} else { //partial selection
				JSONArray ary = (JSONArray) criteria;
				range.enableAutoFilter(field, FilterOp.VALUES, ary.toArray(new String[ary.size()]), null, true);
			}
		}
	}
}
