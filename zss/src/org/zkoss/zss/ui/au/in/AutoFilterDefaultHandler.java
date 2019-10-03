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
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.SortedSet;
import java.util.TreeSet;
import java.util.LinkedHashSet;

import org.zkoss.json.JSONArray;
import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.poi.ss.util.CellReference;
import org.zkoss.util.resource.Labels;
import org.zkoss.zss.api.AreaRef;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.impl.SheetImpl;
import org.zkoss.zss.model.*;
import org.zkoss.zss.model.SFill.FillPattern;
import org.zkoss.zss.model.SAutoFilter.FilterOp;
import org.zkoss.zss.model.SAutoFilter.NFilterColumn;
import org.zkoss.zss.model.SCell.CellType;
import org.zkoss.zss.model.impl.*;
import org.zkoss.zss.model.sys.EngineFactory;
import org.zkoss.zss.model.sys.format.FormatContext;
import org.zkoss.zss.model.sys.format.FormatEngine;
import org.zkoss.zss.model.sys.format.FormatResult;
import org.zkoss.zss.model.util.Strings;
import org.zkoss.zss.model.impl.AbstractAutoFilterAdv.FilterColumnImpl;
import org.zkoss.zss.range.impl.FilterRowInfo;
import org.zkoss.zss.range.impl.FilterRowInfoComparator;
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
		final int index = field - 1; //ZSS-1233
		final NFilterColumn filterColumn = autoFilter.getFilterColumn(index, false);
		String rangeAddr = autoFilter.getRegion().getReferenceString();
		final SRange range = SRanges.range(worksheet, rangeAddr);

		//ZSS-1191
		final SColorFilter colorFilter = filterColumn == null ? null : filterColumn.getColorFilter();
		final SExtraStyle extraStyle = colorFilter == null ? null : colorFilter.getExtraStyle();
		final SFill filterFill = extraStyle == null ? null : extraStyle.getFill();
		final boolean byFontColor = colorFilter == null ? false : colorFilter.isByFontColor();
		
		//ZSS-1192
		final SCustomFilters custFilters = filterColumn == null ? null : filterColumn.getCustomFilters();
		
		//ZSS-1193, ZSS-1234
		final SDynamicFilter dynaFilter = filterColumn == null ? null : filterColumn.getDynamicFilter();
		
		//ZSS-1193
		final STop10Filter top10Filter = filterColumn == null ? null : filterColumn.getTop10Filter();
		
		//ZSS-704: Note that scanRows() could provide new bottom
		final int bottom = range.getLastRow();
		Object[] results = scanRows(field, filterColumn, range, worksheet, 
				table, //ZSS-988 
				filterFill, byFontColor, //ZSS-1191 
				custFilters, //ZSS-1192
				dynaFilter, top10Filter); //ZSS-1193
		@SuppressWarnings("unchecked")
		SortedSet<FilterRowInfo> orderedRowInfos = (SortedSet<FilterRowInfo>) results[0];
		
		if (bottom != ((Integer) results[1]).intValue()) {
			CellRegion region = new CellRegion(range.getRow(), range.getColumn(), bottom, range.getLastColumn()); 
			rangeAddr = region.getReferenceString(); 
		}
		
		//ZSS-1191: 1: Date; 2: Number; 3: Text
		int type = ((Integer)results[3]).intValue();
		
		Set<SFill> ccitems = (Set<SFill>) results[2];
		Set<SFill> fcitems = (Set<SFill>) results[4];
		
		//ZSS-1195
		final String colName = (String) results[5];
		
		//ZSS-1192
		final SCustomFilters customFilters = filterColumn == null ? null : filterColumn.getCustomFilters();
		
		//ZSS-1193
		if (filterColumn != null) {
			((FilterColumnImpl)filterColumn).setFilterType(type);
		}
		
		//ZSS-1233
		((AutoFilterImpl)autoFilter).setCachedSet(index, orderedRowInfos);
		
		//ZSS-1234
		((AutoFilterImpl)autoFilter).setFilterType(index, type);

		spreadsheet.smartUpdate("autoFilterPopup", 
			convertFilterInfoToJSON(row, col, field, rangeAddr, orderedRowInfos,
					type, ccitems, filterFill, fcitems, byFontColor, customFilters,
					dynaFilter, top10Filter, colName)); //ZSS-1191, ZSS-1192, ZSS-1195
		
		AreaRef filterArea = new AreaRef(rangeAddr);
		
		return filterArea;
	}
	
	//see Popup.js#zssex.AutoFilterPopup 
	private Map convertFilterInfoToJSON(int row, int col, int field,
			String rangeAddr, SortedSet<FilterRowInfo> orderedRowInfos,
			int type, Set<SFill> ccitems, SFill filterFill, Set<SFill> fcitems, //ZSS-1191 
			boolean byFontColor, SCustomFilters custFilters, //ZSS-1192
			SDynamicFilter dynaFilter, STop10Filter top10Filter, //ZSS-1193
			String colName) { //ZSS-1195
		final Map data = new HashMap();
		
		boolean selectAll = filterFill == null //ZSS-1191
				&& custFilters == null  //ZSS-1192
				&& dynaFilter == null && top10Filter == null; //ZSS-1193
		boolean select = false;
		final ArrayList<Map> sortedItems = new ArrayList<Map>();
		for (FilterRowInfo info : orderedRowInfos) {			
			if (info == blankRowInfo) {
				data.put("blank", info.isSelected());
				if (info.isSelected()) {
					select = true;
				} else {
					selectAll = false;
				}
			} else {
				HashMap item = new HashMap();
				sortedItems.add(item);
				item.put("v", info.getDisplay());
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
			if (fill.getFillPattern() == FillPattern.NONE && BLANK_FONT.getColor().equals(fill.getFillColor())) {
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
		
		//ZSS-1192
		final SCustomFilter f1 = custFilters == null ? null : custFilters.getCustomFilter1();
		final SCustomFilter f2 = custFilters == null ? null : custFilters.getCustomFilter2();
		FilterOp f1op = null; //ZSS-1234
		FilterOp f2op = null; //ZSS-1234
		if (custFilters != null) {
			final boolean isAnd = custFilters.isAnd();
			data.put("and", isAnd);
			
			final Map f1Map = new HashMap();
			data.put("f1", f1Map);
			f1Map.put("val", f1.getValue());
			f1Map.put("op", (f1op = f1.getOperator()).name()); //ZSS-1234
			
			if (f2 != null) {
				final Map f2Map = new HashMap();
				data.put("f2", f2Map);
				f2Map.put("val", f1.getValue());
				f2Map.put("op", (f2op = f2.getOperator()).name()); //ZSS-1234
			}
		}

		final ArrayList<Map> vitems = new ArrayList<Map>();
		final ArrayList<Map> pitems = new ArrayList<Map>(); //ZSS-1234
		FilterOp targetOp = null;
		FilterOp targetOp2 = null;
		if (custFilters != null) { //ZSS-1192
			if (f1op == FilterOp.greaterThanOrEqual && f2op == FilterOp.lessThanOrEqual && type != 3) {
				f1op = type == 1 ? FilterOp.betweenDates : FilterOp.between;
				f2op = null;
			}
			targetOp = f1op != null && f2op == null ? f1op : 
				f1op != null && f2op != null ? FilterOp.custom : null;
		} else if (dynaFilter != null) { //ZSS-1193, ZSS-1234: SDynamicFilter
			targetOp = FilterOp.valueOf(dynaFilter.getType());
			switch(targetOp) {
			case Q1:
			case Q2:
			case Q3:
			case Q4:
			case M1:
			case M2:
			case M3:
			case M4:
			case M5:
			case M6:
			case M7:
			case M8:
			case M9:
			case M10:
			case M11:
			case M12:
				targetOp2 = targetOp;
				targetOp = FilterOp.allDatesInPeriod;
				break;
			}
		} else if (top10Filter != null) { //ZSS-1193: STop10Filter
			targetOp = FilterOp.top10; 
		}
		if (type == 3) { //ZSS-1192: 1: Date; 2: Number, 3: Text
			opToJson(vitems, FilterOp.equal, targetOp, false);
			opToJson(vitems, FilterOp.notEqual, targetOp, true);
			opToJson(vitems, FilterOp.beginWith, targetOp, false);
			opToJson(vitems, FilterOp.endWith, targetOp, true);
			opToJson(vitems, FilterOp.contains, targetOp, false);
			opToJson(vitems, FilterOp.notContains, targetOp, true);
			opToJson(vitems, FilterOp.custom, targetOp, false);
		} else if (type == 2) { //ZSS-1193: number
			opToJson(vitems, FilterOp.equal, targetOp, false);
			opToJson(vitems, FilterOp.notEqual, targetOp, true);
			opToJson(vitems, FilterOp.greaterThan, targetOp, false);
			opToJson(vitems, FilterOp.greaterThanOrEqual, targetOp, false);
			opToJson(vitems, FilterOp.lessThan, targetOp, false);
			opToJson(vitems, FilterOp.lessThanOrEqual, targetOp, false);
			opToJson(vitems, FilterOp.between, targetOp, true);
			opToJson(vitems, FilterOp.top10, targetOp, false);
			opToJson(vitems, FilterOp.aboveAverage, targetOp, false);
			opToJson(vitems, FilterOp.belowAverage, targetOp, true);
			opToJson(vitems, FilterOp.custom, targetOp, false);
		} else if (type == 1) { //ZSS-1234: Date
			opToJson(vitems, FilterOp.equal, targetOp, true);
			opToJson(vitems, FilterOp.before, targetOp == FilterOp.lessThan ? FilterOp.before : targetOp, false);
			opToJson(vitems, FilterOp.after, targetOp == FilterOp.greaterThan ? FilterOp.after : targetOp, false);
			opToJson(vitems, FilterOp.betweenDates, targetOp, true);
			opToJson(vitems, FilterOp.tomorrow, targetOp, false);
			opToJson(vitems, FilterOp.today, targetOp, false);
			opToJson(vitems, FilterOp.yesterday, targetOp, true);
			opToJson(vitems, FilterOp.nextWeek, targetOp, false);
			opToJson(vitems, FilterOp.thisWeek, targetOp, false);
			opToJson(vitems, FilterOp.lastWeek, targetOp, true);
			opToJson(vitems, FilterOp.nextMonth, targetOp, false);
			opToJson(vitems, FilterOp.thisMonth, targetOp, false);
			opToJson(vitems, FilterOp.lastMonth, targetOp, true);
			opToJson(vitems, FilterOp.nextQuarter, targetOp, false);
			opToJson(vitems, FilterOp.thisQuarter, targetOp, false);
			opToJson(vitems, FilterOp.lastQuarter, targetOp, true);
			opToJson(vitems, FilterOp.nextYear, targetOp, false);
			opToJson(vitems, FilterOp.thisYear, targetOp, false);
			opToJson(vitems, FilterOp.lastYear, targetOp, true);
			opToJson(vitems, FilterOp.yearToDate, targetOp, true);
			opToJson(vitems, FilterOp.allDatesInPeriod, targetOp, true);
			opToJson(vitems, FilterOp.custom, targetOp, false);
			
			opToJson2(pitems, FilterOp.Q1, targetOp2, false);
			opToJson2(pitems, FilterOp.Q2, targetOp2, false);
			opToJson2(pitems, FilterOp.Q3, targetOp2, false);
			opToJson2(pitems, FilterOp.Q4, targetOp2, true);
			opToJson2(pitems, FilterOp.M1, targetOp2, false);
			opToJson2(pitems, FilterOp.M2, targetOp2, false);
			opToJson2(pitems, FilterOp.M3, targetOp2, false);
			opToJson2(pitems, FilterOp.M4, targetOp2, false);
			opToJson2(pitems, FilterOp.M5, targetOp2, false);
			opToJson2(pitems, FilterOp.M6, targetOp2, false);
			opToJson2(pitems, FilterOp.M7, targetOp2, false);
			opToJson2(pitems, FilterOp.M8, targetOp2, false);
			opToJson2(pitems, FilterOp.M9, targetOp2, false);
			opToJson2(pitems, FilterOp.M10, targetOp2, false);
			opToJson2(pitems, FilterOp.M11, targetOp2, false);
			opToJson2(pitems, FilterOp.M12, targetOp2, false);
		}
		data.put("vitems", vitems);
		data.put("vitem", targetOp == null ? null : targetOp.name());

		//ZSS-1234
		if (!pitems.isEmpty()) {
			data.put("pitems", pitems);
			data.put("pitem", targetOp2 == null ? null : targetOp2.name());
		}

		data.put("items", sortedItems);
		data.put("row", row);
		data.put("col", col);
		data.put("field", field);
		data.put("range", rangeAddr);
		data.put("select", selectAll ? "all" : select ? "mix" : "none");
		data.put("colName", colName); //ZSS-1195
		return data;
	}

	//ZSS-1192, ZSS-1234
	private void opToJson(ArrayList<Map> vitems, FilterOp op, FilterOp targetOp, boolean border) {
		HashMap item = new HashMap();
		vitems.add(item);
		item.put("op", op.name());
		item.put("t", Labels.getLabel("zssex.valuedlg."+op.name()));
		if (targetOp == op) {
			item.put("s", true);
		}
		if (border) {
			item.put("b", true);
		}
	}

	//ZSS-1234
	private void opToJson2(ArrayList<Map> vitems, Enum op, Enum targetOp, boolean border) {
		HashMap item = new HashMap();
		vitems.add(item);
		item.put("op", op.name());
		item.put("t", Labels.getLabel("zssex.perdlg."+op.name()));
		if (targetOp == op) {
			item.put("s", true);
		}
		if (border) {
			item.put("b", true);
		}
	}

	// ZSS-704
	// [0]: SortedSet<FilterRowInfo>; 
	// [1]: new bottom; 
	// [2]: LinkedHashSet<SFill> for CELL_COLOR; 
	// [3]: which kind of filter(Number Filter/ Date Filter/ Text Filter);
	// [4]: LinkedHashSet<SFill> for FONT_COLOR
	// [5]: filter column name
	private Object[] scanRows(int field, NFilterColumn fc, SRange range, 
			SSheet worksheet, STable table, //ZSS-988 
			SFill filterFill, boolean byFontColor, //ZSS-1191 
			SCustomFilters custFilters,
			SDynamicFilter dynaFilter, STop10Filter top10Filter) { //ZSS-1192
		SortedSet<FilterRowInfo> orderedRowInfos = 
				new TreeSet<FilterRowInfo>(new FilterRowInfoComparator());
		
		//ZSS-1191
		LinkedHashSet<SFill> ccitems = new LinkedHashSet<SFill>(); //ZSS-1191: CELL_COLOR
		LinkedHashSet<SFill> fcitems = new LinkedHashSet<SFill>(); //ZSS-1191: FONT_COLOR
		int[] types = new int[] {0, 0, 0}; //0: date, 1: number, 2: string
		
		blankRowInfo = new FilterRowInfo(BLANK_VALUE, "(Blanks)");
		final Set criteria1 = fc == null ? null : fc.getCriteria1();
		boolean hasBlank = false;
		boolean hasSelectedBlank = false;
		final int top = range.getRow() + 1;
		int bottom = range.getLastRow();
		final int columnIndex = range.getColumn() + field - 1;
		final SFont defaultFont = worksheet.getBook().getDefaultFont(); //ZSS-1191
		FormatEngine fe = EngineFactory.getInstance().createFormatEngine();
		boolean isItemFilter = filterFill == null  //ZSS-1191 
				&& custFilters == null  //ZSS-1192
				&& dynaFilter == null && top10Filter == null; //ZSS-1193
		//ZSS-1195
		final SCell cell = worksheet.getCell(range.getRow(), columnIndex);
		String colName = null;
		if (!cell.isNull() && cell.getType() != CellType.BLANK) {
			FormatResult fr = fe.format(cell, new FormatContext(ZssContext.getCurrent().getLocale()));
			colName = fr.getText();
		}
		if (colName == null || Strings.isBlank(colName)) {
			final String ab = CellReference.convertNumToColString(columnIndex);
			colName = "(Column "+ab+")";
		}

		for (int i = top; i <= bottom; i++) {
			//ZSS-988: filter column with no criteria should not show option of hidden row 
			if (isItemFilter && (criteria1 == null || criteria1.isEmpty())) { //ZSS-1191, ZSS-1192, ZSS-1193
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
				BLANK_FILL : new FillImpl(FillPattern.NONE, font.getColor(), ColorImpl.WHITE);
			fcitems.add(fcfill);			
			int type0 = 3; //ZSS-1241
			
			if (!c.isNull() && c.getType() != CellType.BLANK) {
				FormatResult fr = fe.format(c, new FormatContext(ZssContext.getCurrent().getLocale()));
				String displaytxt = fr.getText();
				if(!hasBlank && displaytxt.trim().isEmpty()) { //ZSS-707: show as blank; then it is blank
					hasBlank = true;
					hasSelectedBlank = prepareBlankRow(criteria1, hasSelectedBlank, isItemFilter); //ZSS-1191, ZSS-1192, ZSS-1193
				} else {				
					Object val = c.getValue(); // ZSS-707
					if(c.getType()==CellType.NUMBER && fr.isDateFormatted()){
						val = c.getDateValue();
					}
					
					//ZSS-1191: Date 1, Number 2, String 3, Boolean 4 is Number; Error 5 and Blank 6.
					final int type = FilterRowInfo.getType(val);
					type0 = (type == 4 ? 2 : type) - 1; //ZSS-1241
					
					FilterRowInfo rowInfo = new FilterRowInfo(val, displaytxt);
					//ZSS-299
					orderedRowInfos.add(rowInfo);
					//ZSS-1191, ZSS-1192, ZSS-1193: color/custom/dynamic/top10 filter excludes item filter
					if (isItemFilter) { 
						if (criteria1 == null || criteria1.isEmpty() || criteria1.contains(displaytxt)) { //selected
							rowInfo.setSelected(true);
						}
					}
				}
			} else if (!hasBlank){
				hasBlank = true;
				hasSelectedBlank = prepareBlankRow(criteria1, hasSelectedBlank, isItemFilter); //ZSS-1191, ZSS-1192, ZSS-1193
			}
			
			//ZSS-1241: Date 0, Number 1, String 2
			if (type0 < 3) {
				types[type0] = types[type0] + 1;
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
				
				int type0 = 3; //ZSS-1241

				if (!c.isNull() && c.getType() != CellType.BLANK) {
					FormatResult fr = fe.format(c, new FormatContext(ZssContext.getCurrent().getLocale()));
					String displaytxt = fr.getText();
					if(!hasBlank && displaytxt.trim().isEmpty()) { //ZSS-707: show as blank; then it is blank
						hasBlank = true;
						hasSelectedBlank = prepareBlankRow(criteria1, hasSelectedBlank, isItemFilter); //ZSS-1191, ZSS-1192, ZSS-1193
					} else {
						Object val = c.getValue(); // ZSS-707
						if(c.getType()==CellType.NUMBER && fr.isDateFormatted()){
							val = c.getDateValue();
						}
						
						//ZSS-1191: Date 1, Number 2, String 3, Boolean 4 is Number; Error 5 and Blank 6
						final int type = FilterRowInfo.getType(val);
						type0 = (type == 4 ? 2 : type) - 1; // ZSS-1241
						
						FilterRowInfo rowInfo = new FilterRowInfo(val, displaytxt);
						//ZSS-299
						orderedRowInfos.add(rowInfo);
						//ZSS-1191, ZSS-1192: color/custom/dynamic/top10 filter excludes item filter
						if (filterFill == null && custFilters == null) { 
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
					if (!leaveLoop && !hasBlank) { //ZSS-1233
						hasBlank = true;
						hasSelectedBlank = prepareBlankRow(criteria1, hasSelectedBlank, isItemFilter); //ZSS-1191, ZSS-1192, ZSS-1193
					}
				}
				
				if (leaveLoop) {
					break;
				}
				
				//ZSS-1241: Date 0, Number 1, String 2
				if (type0 < 3) {
					types[type0] = types[type0] + 1;
				}
				
				//ZSS-1191
				final SCellStyle style = c.isNull() ? null : c.getCellStyle();
				final SFill ccfill = style == null ? BLANK_FILL : style.getFill();
				ccitems.add(ccfill);
				final SFont font = style == null ? null : style.getFont();
				final SFill fcfill = font == null ? BLANK_FILL : new FillImpl(FillPattern.SOLID, font.getColor(), null);
				fcitems.add(fcfill);
			}
		}
		if (hasBlank) {
			orderedRowInfos.add(blankRowInfo);
		}
		
		//ZSS-1241 determine the candidate
		//which kind of filter (DateFilter/NumberFilter/TextFilter); 
		//0: date, 1: number, 2: string; if same count; Text > Number > Date
		int candidate = 2; 
		int max = types[2];
		if (max < types[1]) {
			candidate = 1;
			max = types[1];
		}
		if (max < types[0]) {
			candidate = 0;
		}
		
		return new Object[] {orderedRowInfos, bottom, 
				ccitems.size() > 1 ? ccitems : Collections.EMPTY_SET, //ZSS-1191 
				new Integer(candidate+1), //ZSS-1192
				fcitems.size() > 1 ? fcitems : Collections.EMPTY_SET, //ZSS-1191
				colName}; //ZSS-1195
	}

	//ZSS-707
	private boolean prepareBlankRow(Set criteria1, boolean hasSelectedBlank, 
			boolean isItemFilter) { //ZSS-1191, ZSS-1192, ZSS-1193
		boolean noFilterApplied = criteria1 == null || criteria1.isEmpty();
		//ZSS-1911, ZSS-1192, ZSS-1193: color/custom/dynamic/top10 filter excludes item filter
		if (isItemFilter) { 
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
	 
	/*package*/ void applyFilter(Spreadsheet spreadsheet, Sheet selectedSheet,
			String cellRangeAddr, boolean selectAll, int field, Object criteria, FilterOp op) { //ZSS-1191
		final SRange range = SRanges.range(((SheetImpl)selectedSheet).getNative(), cellRangeAddr);

		//ZSS-1234
		switch(op) {
		//ZSS-1191: criteria: [pattern, fgcolor, bgcolor]
		case cellColor:
		case fontColor:
			
		//ZSS-1192: criteria: [SCustomFilter.Operator1, val1, SCustomFilter.Operator2, val2]
		case and:
		case or:
		{
			JSONArray ary = (JSONArray) criteria;
			range.enableAutoFilter(field, op, ary.toArray(new String[ary.size()]), null, true);
			break;
		}
		
		//ZSS-1193
		case aboveAverage:
		case belowAverage:
			
		//ZSS-1234
		case tomorrow:
		case today:
		case yesterday:
		case nextWeek:
		case thisWeek:
		case lastWeek:
		case nextMonth:
		case thisMonth:
		case lastMonth:
		case nextQuarter:
		case thisQuarter:
		case lastQuarter:
		case nextYear:
		case thisYear:
		case lastYear:
		case yearToDate:
		case Q1:
		case Q2:
		case Q3:
		case Q4:
		case M1:
		case M2:
		case M3:
		case M4:
		case M5:
		case M6:
		case M7:
		case M8:
		case M9:
		case M10:
		case M11:
		case M12:
			range.enableAutoFilter(field, op, null, null, true);
			break;
			
		default:
			if (selectAll) {
				range.enableAutoFilter(field, FilterOp.values, null, null, true);
			} else { //partial selection
				JSONArray ary = (JSONArray) criteria;
				range.enableAutoFilter(field, FilterOp.values, ary.toArray(new String[ary.size()]), null, true);
			}
		}
	}
}
