/* Consts.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 15, 2010 9:59:50 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app;

/**
 * @author Sam
 *
 */
public class Consts {

	public static String _CellStylePanel_zul = "~./zssapp/html/cellStylePanel.zul";
	
	/*File zul*/
	public static String _FileListOpen_zul = "~./zssapp/html/dialog/file/fileListOpen.zul";
	public static String _ImportFile_zul = "~./zssapp/html/dialog/file/importFile.zul";
	
	/*Hyperlink zull*/
	public static String _Weblink_zul = "~./zssapp/html/dialog/hyperlink/webLink.zul";
	public static String _Doclink_zul = "~./zssapp/html/dialog/hyperlink/docLink.zul";
	public static String _Maillink_zul = "~./zssapp/html/dialog/hyperlink/mailLink.zul";
	
	/*Menu zul*/
	public static String _FormatMenu_zul = "~./zssapp/html/menu/formatMenu.zul";
	public static String _RowHeaderMenu_zul = "~./zssapp/html/menu/rowHeaderMenupopup.zul";
	public static String _ColumnHeaderMenu_zul = "~./zssapp/html/menu/columnHeaderMenupopup.zul";
	public static String _CellMenupopup_zul = "~./zssapp/html/menu/cellMenupopup.zul";
	public static String _CellContext_zul = "~./zssapp/html/cellContext.zul";
	
	/*Dialog zul*/
	public static String _InsertFormulaDialog_zul = "~./zssapp/html/dialog/insertFormula.zul";
	public static String _InsertHyperlinkDialog_zul = "~./zssapp/html/dialog/hyperlink/insertHyperlink.zul";
	public static String _PasteSpecialDialog_zul = "~./zssapp/html/dialog/pasteSpecialWindow.zul";


	/*Style event*/
	public static final String ON_STYLING_TARGET_CHANGED = "onStylingTargetChanged";
	public static final String ON_CELL_STYLE_CHANGED = "onCellStyleChanged";
	
	/*Sheet event*/
	public static final String ON_SHEET_REGAIN_FOCUS = "onSheetRegainFocus";
	public static final String ON_SHEET_REFRESH = "onSheetRefresh";
	public static final String ON_SHEET_OPEN = "onSheetOpen";
	public static final String ON_SHEET_RENAME = "onSheetRename";
	public static final String ON_SHEET_SELECT = "onSheetSelect";
	public static final String ON_SHEET_INSERT_FORMULZ_DIALOG = "onSheetInsertFormulzDialog";
	public static final String ON_SHEET_EXPORT_PDF_DIALOG = "onSheetExportPDFDialog";
	public static final String ON_SHEET_INSERT = "onSheetInsert";
	public static final String ON_SHEET_INSERT_IMAGE = "onSheetInsertImage";
	public static final String ON_SHEET_CUT_SELECTION = "onSheetCutSelection";
	public static final String ON_SHEET_COPY_SELECTION = "onSheetCopySelection";
	public static final String ON_SHEET_PASTE_SELECTION = "onSheetPasteSelection";
	public static final String ON_SHEET_PASTE_SPECIAL_DIALOG = "onSheetPasteSpecialDialog";
	public static final String ON_SHEET_CLEAR_SELECTION_CONTENT = "onSheetClearSelectionContent";
	public static final String ON_SHEET_CLEAR_SELECTION_STYLE = "onSheetClearSelectionStyle";
	public static final String ON_SHEET_INSERT_ROW = "onSheetInsertRow";
	public static final String ON_SHEET_DELETE_ROW = "onSheetDeleteRow";
	public static final String ON_SHEET_INSERT_COLUMN = "onSheetInsertColumn";
	public static final String ON_SHEET_DELETE_COLUMN = "onSheetDeleteColumn";
	public static final String ON_SHEET_MODIFY_ROW_HEIGHT_DIALOG = "onSheetModifyRowHeightDialog";
	public static final String ON_SHEET_HIDE = "onSheetHide";
	public static final String ON_SHEET_CUSTOM_SORT_DIALOG = "onSheetCustomSortDialog";
	public static final String ON_SHEET_SORT = "onSheetSort";
	public static final String ON_SHEET_SHIFT_CELL = "onSheetShiftCell";
	public static final String ON_SHEET_MERGE_CELL = "onSheetMergeCell";
	public static final String ON_SHEET_HYPERLINK_DIALOG = "onSheetHyperlinkDialog";
	public static final String ON_SHEET_INSERT_FORMULA = "onSheetInsertFormula";
	
	/*Resource event*/
	public static final String ON_RESOURCE_OPEN_NEW = "onResourceOpenNew";
	public static final String ON_RESOURCE_OPEN = "onResourceOpen";
}