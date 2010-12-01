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
	public static String _EditMenu_zul = "~./zssapp/html/menu/editMenu.zul";
	public static String _FileMenu_zul = "~./zssapp/html/menu/fileMenu.zul";
	public static String _ViewMenu_zul = "~./zssapp/html/menu/viewMenu.zul";
	public static String _InsertMenu_zul = "~./zssapp/html/menu/insertMenu.zul";
	
	
	/*Dialog zul*/
	public static String _InsertFormulaDialog_zul = "~./zssapp/html/dialog/insertFormula.zul";
	public static String _InsertFormulaDialog2_zul = "~./zssapp/html/dialog/insertFormulaDialog.zul";
	public static String _InsertHyperlinkDialog_zul = "~./zssapp/html/dialog/hyperlink/insertHyperlink.zul";
	public static String _PasteSpecialDialog_zul = "~./zssapp/html/dialog/pasteSpecialWindow.zul";
	public static String _ComposeFormulaDialog_zul = "~./zssapp/html/dialog/composeFormulaDialog.zul";
	public static String _FormatNumberDialog_zul = "~./zssapp/html/dialog/formatNumber.zul";

	/* Key */
	public static String KEY_ARG_FORMULA_METAINFO = "KEY_ARG_FORMULA_METAINFO";

	/*Style event*/
	public static final String ON_STYLING_TARGET_CHANGED = "onStylingTargetChanged";
	public static final String ON_CELL_STYLE_CHANGED = "onCellStyleChanged";
	
	/*Sheet event*/
	public static final String ON_SHEET_REFRESH = "onSheetRefresh";
	/**
	 * Fire when spreadsheet set new book or close book
	 */
	public static final String ON_SHEET_OPEN = "onSheetOpen";
	/**
	 * Fired when selected sheet changed
	 */
	public static final String ON_SHEET_CHANGED = "onSheetChanged";
	public static final String ON_SHEET_MERGE_CELL = "onSheetMergeCell";
	public static final String ON_SHEET_INSERT_FORMULA = "onSheetInsertFormula";
	
	/*Resource event*/
	public static final String ON_RESOURCE_OPEN_NEW = "onResourceOpenNew";
	public static final String ON_RESOURCE_OPEN = "onResourceOpen";

	
}