/* Toolbars.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Feb 24, 2012 11:57:44 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ui;

import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;

/**
 * @author sam
 *
 */
public enum Action {
	
	SHEET("sheet"),
	ADD_SHEET("addSheet"),
	DELETE_SHEET("deleteSheet"),
	RENAME_SHEET("renameSheet"),
	MOVE_SHEET_LEFT("moveSheetLeft"),
	MOVE_SHEET_RIGHT("moveSheetRight"),
	PROTECT_SHEET("protectSheet"),
	GRIDLINES("gridlines"),
	HOME_PANEL("homePanel"),
	FORMULA_PANEL("formulaPanel"),
	INSERT_PANEL("insertPanel"),
	NEW_BOOK("newBook"),
	SAVE_BOOK("saveBook"),
	CLOSE_BOOK("closeBook"),
	EXPORT_PDF("exportPDF"),
	PASTE("paste"),
	PASTE_FORMULA("pasteFormula"),
	PASTE_VALUE("pasteValue"),
	PASTE_ALL_EXPECT_BORDERS("pasteAllExceptBorder"),
	PASTE_TRANSPOSE("pasteTranspose"),
	PASTE_SPECIAL("pasteSpecial"),
	CUT("cut"),
	COPY("copy"),
	FONT_FAMILY("fontFamily"),
	FONT_SIZE("fontSize"),
	FONT_BOLD("fontBold"),
	FONT_ITALIC("fontItalic"),
	FONT_UNDERLINE("fontUnderline"),
	FONT_STRIKE("fontStrike"),
	BORDER("border"),
	BORDER_BOTTOM("borderBottom"),
	BORDER_TOP("borderTop"),
	BORDER_LEFT("borderLeft"),
	BORDER_RIGHT("borderRight"),
	BORDER_NO("borderNo"),
	BORDER_ALL("borderAll"),
	BORDER_OUTSIDE("borderOutside"),
	BORDER_INSIDE("borderInside"),
	BORDER_INSIDE_HORIZONTAL("borderInsideHorizontal"),
	BORDER_INSIDE_VERTICAL("borderInsideVertical"),
	BORDER_COLOR("borderColor"),
	FILL_COLOR("fillColor"),
	FONT_COLOR("fontColor"),
	VERTICAL_ALIGN("verticalAlign"),
	VERTICAL_ALIGN_TOP("verticalAlignTop"),
	VERTICAL_ALIGN_MIDDLE("verticalAlignMiddle"),
	VERTICAL_ALIGN_BOTTOM("verticalAlignBottom"),
	HORIZONTAL_ALIGN("horizontalAlign"),
	HORIZONTAL_ALIGN_LEFT("horizontalAlignLeft"),
	HORIZONTAL_ALIGN_CENTER("horizontalAlignCenter"),
	HORIZONTAL_ALIGN_RIGHT("horizontalAlignRight"),
	WRAP_TEXT("wrapText"),
	MERGE_AND_CENTER("mergeAndCenter"),
	MERGE_ACROSS("mergeAcross"),
	MERGE_CELL("mergeCell"),
	UNMERGE_CELL("unmergeCell"),
	INSERT("insert"),
	INSERT_CELL("insertCell"),
	INSERT_SHIFT_CELL_RIGHT("shiftCellRight"),
	INSERT_SHIFT_CELL_DOWN("shiftCellDown"),
	INSERT_SHEET_ROW("insertSheetRow"),
	INSERT_SHEET_COLUMN("insertSheetColumn"),
	DELETE("del"),
	DELETE_CELL("deleteCell"),
	DELETE_SHIFT_CELL_LEFT("shiftCellLeft"),
	DELETE_SHIFT_CELL_UP("shiftCellUp"),
	DELETE_SHEET_ROW("deleteSheetRow"),
	DELETE_SHEET_COLUMN("deleteSheetColumn"),
	FORMAT("format"),
	ROW_HEIGHT("rowHeight"),
	COLUMN_WIDTH("columnWidth"),
	HIDE_ROW("hideRow"),
	UNHIDE_ROW("unhideRow"),
	HIDE_COLUMN("hideColumn"),
	UNHIDE_COLUMN("unhideColumn"),
	FORMAT_CELL("formatCell"),
	LOCK_CELL("lockCell"),
	CLEAR("clear"),
	CLEAR_CONTENT("clearContent"),
	CLEAR_STYLE("clearStyle"),
	CLEAR_ALL("clearAll"),
	AUTO_SUM("autoSum"),
	COUNT_NUMBER("countNumber"),
	MAX("max"),
	MIN("min"),
	SORT("sort"),
	SORT_AND_FILTER("sortAndFilter"),
	SORT_ASCENDING("sortAscending"),
	SORT_DESCENDING("sortDescending"),
	CUSTOM_SORT("customSort"),
	FILTER("filter"),
	CLEAR_FILTER("clearFilter"),
	REAPPLY_FILTER("reapplyFilter"),
	INSERT_PICTURE("insertPicture"),
	COLUMN_CHART("columnChart"),
	COLUMN_CHART_3D("columnChart3D"),
	LINE_CHART("lineChart"),
	LINE_CHART_3D("lineChart3D"),
	PIE_CHART("pieChart"),
	PIE_CHART_3D("pieChart3D"),
	BAR_CHART("barChart"),
	BAR_CHART_3D("barChart3D"),
	AREA_CHART("areaChart"),
	SCATTER_CHART("scatterChart"),
	OTHER_CHART("otherChart"),
	DOUGHNUT_CHART("doughnutChart"),
	HYPERLINK("hyperlink"),
	INSERT_FUNCTION("insertFunction"),
	FINANCIAL("financial"),
	LOGICAL("logical"),
	TEXT("Text"),
	DATE_AND_TIME("dateAndTime"),
	LOOKUP_AND_REFERENCE("lookupAndReference"),
	MATH_AND_TRIG("mathAndTrig"),
	MORE_FUNCTION("moreFunction"),
	/**
	 * A marker stand for toolbarbutton separator symbol (for client side only)
	 */
	SEPARATOR("separator");
	
	private final String action;
	private Action(String key) {
		this.action = key;
	}
	
	public String getLabelKey() {
		return "zss." + action;
	}
	
	public boolean equals(String action) {
		return this.action.equals(action);
	}
	
	@Override
	public String toString() {
		return action;
	}
	
	public static Collection<String> getLabelKeys() {
		Action[] enums = Action.class.getEnumConstants();
		ArrayList<String> keys = new ArrayList<String>(enums.length);
		for (Action a : enums) {
			keys.add(a.getLabelKey());
		}
		return keys;
	}
	
	/**
	 * Returns all Action
	 * 
	 * @return
	 */
	public static HashMap<String, Action> getAll() {
		Action[] enums = Action.class.getEnumConstants();
		HashMap<String, Action> actions = new HashMap<String, Action>(enums.length);
		for (Action t : enums) {
			actions.put(t.toString(), t);
		}
		return actions;
	}
}
