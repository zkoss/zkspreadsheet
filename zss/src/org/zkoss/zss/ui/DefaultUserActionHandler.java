/* DefaultUserActionHandler.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/06/03 by Dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
 */
package org.zkoss.zss.ui;

import java.util.Arrays;
import java.util.LinkedHashSet;
import java.util.Map;
import java.util.Set;

import org.zkoss.lang.Strings;
import org.zkoss.util.resource.Labels;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.ApplyBorderType;
import org.zkoss.zss.api.Range.AutoFillType;
import org.zkoss.zss.api.Range.DeleteShift;
import org.zkoss.zss.api.Range.InsertCopyOrigin;
import org.zkoss.zss.api.Range.InsertShift;
import org.zkoss.zss.api.Range.PasteOperation;
import org.zkoss.zss.api.Range.PasteType;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.SheetOperationUtil;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.CellStyle.Alignment;
import org.zkoss.zss.api.model.CellStyle.BorderType;
import org.zkoss.zss.api.model.CellStyle.VerticalAlignment;
import org.zkoss.zss.api.model.Font.Boldweight;
import org.zkoss.zss.api.model.Font.Underline;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.DefaultUserActionHandler.Clipboard.Type;
import org.zkoss.zss.ui.event.AuxActionEvent;
import org.zkoss.zss.ui.event.CellSelectionAction;
import org.zkoss.zss.ui.event.Events;
import org.zkoss.zss.ui.event.KeyEvent;
import org.zkoss.zss.ui.event.CellSelectionUpdateEvent;
import org.zkoss.zul.Messagebox;
/**
 * The user action handler which provide default spreadsheet operation handling.
 *  
 * @author dennis
 * @since 3.0.0
 */
public class DefaultUserActionHandler implements UserActionHandler {
	
	private static final long serialVersionUID = 1L;
	private static ThreadLocal<UserActionContext> _ctx = new ThreadLocal<UserActionContext>();
	private Clipboard _clipboard;
	
	protected void checkCtx(){
		if(_ctx.get()==null){
			throw new IllegalAccessError("can't found action context");
		}
	}
	
	/**
	 * Gets the spreadsheet of current action
	 * @return
	 */
	protected Spreadsheet getSpreadsheet(){
		checkCtx();
		return _ctx.get().spreadsheet;
	}
	
	/**
	 * Gets the sheet of current action
	 * @return
	 */
	protected Sheet getSheet(){
		checkCtx();
		return _ctx.get().sheet;
	}
	
	/**
	 * Gets the book of current action
	 * @return
	 */
	protected Book getBook(){
		checkCtx();
		Sheet sheet = getSheet();
		return sheet==null?null:sheet.getBook();
	}
	
	  /**
	  * Gets the selection of current action, it is always smaller than
	  * {@link Spreadsheet#getMaxVisibleRows()} and {@link Spreadsheet#getMaxVisibleColumns()} of spreadsheet selection. <br/> 
	  * Note: gets selection from {@link Spreadsheet#getSelection()} will get the last user
	  * selection which might be a whole sheet/row/column selection.
	  */
	protected Rect getSelection(){
		checkCtx();
		return _ctx.get().selection;
	}
	
	protected Object getExtraData(String key){
		checkCtx();
		Map data = _ctx.get().extraData;
		if(data!=null){
			return data.get(key);
		}
		return null;
	}
	
	protected boolean dispatchAction(String action) {
		
		DefaultUserAction dua = DefaultUserAction.getBy(action);
		if(dua==null){
			return doCustom(action);
		}
		
		
		/*if (UserAction.HOME_PANEL.equals(dua)) {
			return doShowHomePanel();
		} else if (UserAction.INSERT_PANEL.equals(dua)) {
			return doShowInsertPanel();
		} else if (UserAction.FORMULA_PANEL.equals(dua)) {
			return doShowFormulaPanel();
		} else*/if (DefaultUserAction.CLOSE_BOOK.equals(dua)) {
			return doCloseBook();
		}else if (DefaultUserAction.ADD_SHEET.equals(dua)) {
			return doAddSheet();
		} else if (DefaultUserAction.DELETE_SHEET.equals(dua)) {
			return doDeleteSheet();
		} else if (DefaultUserAction.RENAME_SHEET.equals(dua)) {
			String name = (String) getExtraData("name");
			return doRenameSheet(name);
		} else if (DefaultUserAction.MOVE_SHEET_LEFT.equals(dua)) {
			return doMoveSheetLeft();
		} else if (DefaultUserAction.MOVE_SHEET_RIGHT.equals(dua)) {
			return doMoveSheetRight();
		} else if (DefaultUserAction.PROTECT_SHEET.equals(dua)) {
			return doProtectSheet();
		} else if (DefaultUserAction.GRIDLINES.equals(dua)) {
			return doGridlines();
		} /*else if (UserAction.NEW_BOOK.equals(dua)) {
			return doNewBook();
		} else if (UserAction.SAVE_BOOK.equals(dua)) {
			return doSaveBook();
		} else if (UserAction.EXPORT_PDF.equals(dua)) {
			return doExportPDF();
		} */else if (DefaultUserAction.PASTE.equals(dua)) {
			return doPaste();
		} else if (DefaultUserAction.PASTE_FORMULA.equals(dua)) {
			return doPasteFormula();
		} else if (DefaultUserAction.PASTE_VALUE.equals(dua)) {
			return doPasteValue();
		} else if (DefaultUserAction.PASTE_ALL_EXPECT_BORDERS.equals(dua)) {
			return doPasteAllExceptBorder();
		} else if (DefaultUserAction.PASTE_TRANSPOSE.equals(dua)) {
			return doPasteTranspose();
		} /*else if (UserAction.PASTE_SPECIAL.equals(dua)) {
			return doPasteSpecial();
		} */else if (DefaultUserAction.CUT.equals(dua)) {
			return doCut();	
		} else if (DefaultUserAction.COPY.equals(dua)) {
			return doCopy();
		} else if (DefaultUserAction.FONT_FAMILY.equals(dua)) {
			String name = (String) getExtraData("name");
			return doFontFamily(name);
		} else if (DefaultUserAction.FONT_SIZE.equals(dua)) {
			Integer fontSize = Integer.parseInt((String)getExtraData("size"));
			return doFontSize(fontSize);
		} else if (DefaultUserAction.FONT_BOLD.equals(dua)) {
			return doFontBold();
		} else if (DefaultUserAction.FONT_ITALIC.equals(dua)) { 
			return doFontItalic();
		} else if (DefaultUserAction.FONT_UNDERLINE.equals(dua)) {
			return doFontUnderline();
		} else if (DefaultUserAction.FONT_STRIKE.equals(dua)) {
			return doFontStrikeout();
		} else if (DefaultUserAction.BORDER.equals(dua)) {
			return doBorder(getBorderColor());
		} else if (DefaultUserAction.BORDER_BOTTOM.equals(dua)) {
			return doBorderBottom(getBorderColor());
		} else if (DefaultUserAction.BORDER_TOP.equals(dua)) {
			return doBoderTop(getBorderColor());
		} else if (DefaultUserAction.BORDER_LEFT.equals(dua)) {
			return doBorderLeft(getBorderColor());
		} else if (DefaultUserAction.BORDER_RIGHT.equals(dua)) {
			return doBorderRight(getBorderColor());
		} else if (DefaultUserAction.BORDER_NO.equals(dua)) {
			return doBorderNo(getBorderColor());
		} else if (DefaultUserAction.BORDER_ALL.equals(dua)) {
			return doBorderAll(getBorderColor());
		} else if (DefaultUserAction.BORDER_OUTSIDE.equals(dua)) {
			return doBorderOutside(getBorderColor());
		} else if (DefaultUserAction.BORDER_INSIDE.equals(dua)) {
			return doBorderInside(getBorderColor());
		} else if (DefaultUserAction.BORDER_INSIDE_HORIZONTAL.equals(dua)) {
			return doBorderInsideHorizontal(getBorderColor());
		} else if (DefaultUserAction.BORDER_INSIDE_VERTICAL.equals(dua)) {
			return doBorderInsideVertical(getBorderColor());
		} else if (DefaultUserAction.FONT_COLOR.equals(dua)) {
			return doFontColor(getFontColor());
		} else if (DefaultUserAction.FILL_COLOR.equals(dua)) {
			return doFillColor(getFillColor());
		} else if (DefaultUserAction.VERTICAL_ALIGN_TOP.equals(dua)) {
			return doVerticalAlignTop();
		} else if (DefaultUserAction.VERTICAL_ALIGN_MIDDLE.equals(dua)) {
			return doVerticalAlignMiddle();
		} else if (DefaultUserAction.VERTICAL_ALIGN_BOTTOM.equals(dua)) {
			return doVerticalAlignBottom();
		} else if (DefaultUserAction.HORIZONTAL_ALIGN_LEFT.equals(dua)) {
			return doHorizontalAlignLeft();
		} else if (DefaultUserAction.HORIZONTAL_ALIGN_CENTER.equals(dua)) {
			return doHorizontalAlignCenter();
		} else if (DefaultUserAction.HORIZONTAL_ALIGN_RIGHT.equals(dua)) {
			return doHorizontalAlignRight();
		} else if (DefaultUserAction.WRAP_TEXT.equals(dua)) {
			return doWrapText();
		} else if (DefaultUserAction.MERGE_AND_CENTER.equals(dua)) {
			return doMergeAndCenter();
		} else if (DefaultUserAction.MERGE_ACROSS.equals(dua)) {
			return doMergeAcross();
		} else if (DefaultUserAction.MERGE_CELL.equals(dua)) {
			return doMergeCell();
		} else if (DefaultUserAction.UNMERGE_CELL.equals(dua)) {
			return doUnmergeCell();
		} else if (DefaultUserAction.INSERT_SHIFT_CELL_RIGHT.equals(dua)) {
			return doShiftCellRight();
		} else if (DefaultUserAction.INSERT_SHIFT_CELL_DOWN.equals(dua)) {
			return doShiftCellDown();
		} else if (DefaultUserAction.INSERT_SHEET_ROW.equals(dua)) {
			return doInsertSheetRow();
		} else if (DefaultUserAction.INSERT_SHEET_COLUMN.equals(dua)) {
			return doInsertSheetColumn();
		} else if (DefaultUserAction.DELETE_SHIFT_CELL_LEFT.equals(dua)) {
			return doShiftCellLeft();
		} else if (DefaultUserAction.DELETE_SHIFT_CELL_UP.equals(dua)) {
			return doShiftCellUp();
		} else if (DefaultUserAction.DELETE_SHEET_ROW.equals(dua)) {
			return doDeleteSheetRow();
		} else if (DefaultUserAction.DELETE_SHEET_COLUMN.equals(dua)) {
			return doDeleteSheetColumn();
		} else if (DefaultUserAction.SORT_ASCENDING.equals(dua)) {
			return doSortAscending();
		} else if (DefaultUserAction.SORT_DESCENDING.equals(dua)) {
			return doSortDescending();
		} /*else if (UserAction.CUSTOM_SORT.equals(dua)) {
			return doCustomSort();
		} */else if (DefaultUserAction.FILTER.equals(dua)) {
			return doFilter();
		} else if (DefaultUserAction.CLEAR_FILTER.equals(dua)) {
			return doClearFilter();
		} else if (DefaultUserAction.REAPPLY_FILTER.equals(dua)) {
			return doReapplyFilter();
		} else if (DefaultUserAction.CLEAR_CONTENT.equals(dua)) {
			return doClearContent();
		} else if (DefaultUserAction.CLEAR_STYLE.equals(dua)) {
			return doClearStyle();
		} else if (DefaultUserAction.CLEAR_ALL.equals(dua)) {
			return doClearAll();
		} /*else if (UserAction.HYPERLINK.equals(dua)) {
			return doHyperlink();
		} else if (UserAction.FORMAT_CELL.equals(dua)) {
			return doFormatCell();
		} else if (UserAction.COLUMN_WIDTH.equals(dua)) {
			return doColumnWidth();
		} else if (UserAction.ROW_HEIGHT.equals(dua)) {
			return doRowHeight();
		} */else if (DefaultUserAction.HIDE_COLUMN.equals(dua)) {
			return doHideColumn();
		} else if (DefaultUserAction.UNHIDE_COLUMN.equals(dua)) {
			return doUnhideColumn();
		} else if (DefaultUserAction.HIDE_ROW.equals(dua)) {
			return doHideRow();
		} else if (DefaultUserAction.UNHIDE_ROW.equals(dua)) {
			return doUnhideRow();
		} /*else if (UserAction.INSERT_FUNCTION.equals(dua)) {
			return doInsertFunction();
		} */else{
			return doCustom(action);
		}
		
		
	}
	
	protected boolean doCustom(String action){
		showNotImplement(action);
		return false;
	}

	protected boolean doMoveSheetRight() {
		Book book = getBook();
		Sheet sheet = getSheet();
		
		int max = book.getNumberOfSheets();
		int i = book.getSheetIndex(sheet);
		
		if(i<max){
			i ++;
			Range range = Ranges.range(sheet);
			SheetOperationUtil.setSheetOrder(range,i);
		}
		return true;
	}

	protected boolean doMoveSheetLeft() {
		Book book = getBook();
		Sheet sheet = getSheet();
		
		int i = book.getSheetIndex(sheet);
		
		if(i>0){
			i --;
			Range range = Ranges.range(sheet);
			SheetOperationUtil.setSheetOrder(range,i);
		}
		return true;
	}

	protected boolean doRenameSheet(String newname) {
		Book book = getBook();
		Sheet sheet = getSheet();
		
		if(sheet.getSheetName().equals(newname)){
			return true;
		}
		if(!isLeaglSheetName(newname)){
			showWarnMessage("The name :"+newname+", is not a legal sheet name");
			return true;
		}
		if(book.getSheet(newname)!=null){
			showWarnMessage("Canot rename a sheet to the same as another.");
			return true;
		}
		
		Range range = Ranges.range(sheet);
		SheetOperationUtil.renameSheet(range,newname);
		return true;
	}

	protected boolean isLeaglSheetName(String newname) {
		if(Strings.isEmpty(newname))
			return false;
		
		if(newname.length()>31){
			return false;
		}
		
		// \/?*[] are illegal
		String regx = ".*[\\\\\\/\\?\\*\\[\\]]+.*";
		if(newname.matches(regx)){
			return false;
		}
		
		return true;
	}

	protected boolean doDeleteSheet() {
		Book book = getBook();
		Sheet sheet = getSheet();
		
		int num = book.getNumberOfSheets();
		if(num<=1){
			showWarnMessage("Canot delete last sheet.");
			return true;
		}
		
		int index = book.getSheetIndex(sheet);
		
		Range range = Ranges.range(sheet);
		SheetOperationUtil.deleteSheet(range);
		
		if(index==num-1){
			index--;
		}
		
		getSpreadsheet().setSelectedSheet(book.getSheetAt(index).getSheetName());
		
		return true;
	}

	protected boolean doAddSheet() {
		String prefix = Labels.getLabel("zss.newSheetPrefix");//TODO define somewhere
		if (Strings.isEmpty(prefix))
			prefix = "Sheet";
		Sheet sheet = getSheet();

		Range range = Ranges.range(sheet);
		SheetOperationUtil.addSheet(range,prefix);
		
		return true;
	}

	protected String getDefaultBorderColor() {
		return "#000000";
	}
	
	protected String getDefaultFontColor() {
		return "#000000";
	}
	
	protected String getDefaultFillColor() {
		return "#FFFFFF";
	}
	
	private String getBorderColor(){
		String color = (String)getExtraData("color");
		if (Strings.isEmpty(color)) {//CE version won't provide color
			color = getDefaultBorderColor();
		}
		return color;
	}
	
	private String getFontColor(){
		String color = (String)getExtraData("color");
		if (Strings.isEmpty(color)) {//CE version won't provide color
			color = getDefaultFontColor();
		}
		return color;
	}
	
	private String getFillColor(){
		String color = (String)getExtraData("color");
		if (Strings.isEmpty(color)) {//CE version won't provide color
			color = getDefaultFillColor();
		}
		return color;
	}
	
	protected boolean doHideRow() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		range = range.toRowRange();
		CellOperationUtil.hide(range);
		return true;
	}

	protected boolean doUnhideRow() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		range = range.toRowRange();
		CellOperationUtil.unHide(range);
		return true;
	}

	protected boolean doUnhideColumn() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		range = range.toColumnRange();
		CellOperationUtil.unHide(range);
		
		return true;
		
	}

	protected boolean doHideColumn() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		range = range.toColumnRange();
		CellOperationUtil.hide(range);
		return true;
	}
	
	protected boolean doCloseBook() {
		Spreadsheet zss = getSpreadsheet();
		if(zss.getSrc()!=null){
			zss.setSrc(null);
		}
		if(zss.getBook()!=null){
			zss.setBook(null);
		}
		
		clearClipboard();
		return true;
	}
	
	protected boolean doKeystroke(int keyCode,boolean ctrlKey, boolean shiftKey, boolean altKey) {
		if (46 == keyCode) {
			if (ctrlKey)
				return doClearStyle();
			else
				return doClearContent();
		}
		if (!ctrlKey)
			return false;
		
		switch (keyCode) {
		case 'X':
			return doCut();
		case 'C':
			return doCopy();
		case 'V':
			return doPaste();
		case 'D':
			return doClearContent();
		case 'B':
			return doFontBold();
		case 'I':
			return doFontItalic();
		case 'U':
			return doFontUnderline();
		}
		return false;
	}
	
	protected Clipboard getClipboard() {
		return _clipboard;
	}
	
	protected void setClipboard(Clipboard clipboard){
		_clipboard = clipboard;
		if(_clipboard!=null && _ctx.get()!=null)
			getSpreadsheet().setHighlight(_clipboard.sourceRect);
	}
	
	protected void clearClipboard() {
		_clipboard = null;
		if(_ctx.get()!=null)
			getSpreadsheet().setHighlight(null);
	}
	
	protected boolean doCopy() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		setClipboard(new Clipboard(Clipboard.Type.COPY, sheet.getSheetName(), selection));
		return true;
	}
	
	protected boolean doPaste(PasteType pasteType, PasteOperation pasteOperation, boolean skipBlank, boolean transpose) {
		Clipboard cb = getClipboard();
		if(cb==null)
			return false;
		
		Book book = getBook();
		Sheet destSheet = getSheet();
		Sheet srcSheet = book.getSheet(cb.sourceSheetName);
		if(srcSheet==null){
			//TODO message;
			clearClipboard();
			return true;
		}
		Rect src = cb.sourceRect;
		
		Rect selection = getSelection();
		
		Range srcRange = Ranges.range(srcSheet, src.getTop(),
				src.getLeft(), src.getBottom(),src.getRight());

		Range destRange = Ranges.range(destSheet, selection.getTop(),
				selection.getLeft(), selection.getBottom(), selection.getRight());
		
		if (destRange.isProtected()) {
			showProtectMessage();
			return true;
		} else if (cb.type == Type.CUT && srcRange.isProtected()) {
			showProtectMessage();
			return true;
		}
		
		if(cb.type==Type.CUT){
			CellOperationUtil.cut(srcRange,destRange);
			clearClipboard();
		}else{
			CellOperationUtil.pasteSpecial(srcRange, destRange, pasteType, pasteOperation, skipBlank, transpose);
		}
		return true;
	}
	
	protected void showProtectMessage() {
		Messagebox.show("The cell that you are trying to change is protected and locked.", "ZK Spreadsheet", Messagebox.OK, Messagebox.EXCLAMATION);
	}
	
	protected void showWarnMessage(String message) {
		Messagebox.show(message, "ZK Spreadsheet", Messagebox.OK, Messagebox.EXCLAMATION);
	}
	
	protected void showNotImplement(String action) {
		Messagebox.show("This action "+Labels.getLabel(action,action)+" doesn't be implemented yet", "ZK Spreadsheet", Messagebox.OK, Messagebox.EXCLAMATION);
	}
	
	protected boolean doPaste() {
		return doPaste(PasteType.ALL,PasteOperation.NONE,false,false);
	}
	
	protected boolean doPasteFormula() {
		return doPaste(PasteType.FORMULAS,PasteOperation.NONE,false,false);
	}
	
	protected boolean doPasteValue() {
		return doPaste(PasteType.VALUES,PasteOperation.NONE,false,false);
	}
	
	protected boolean doPasteAllExceptBorder() {
		return doPaste(PasteType.ALL_EXCEPT_BORDERS,PasteOperation.NONE,false,false);
	}
	
	protected boolean doPasteTranspose() {
		return doPaste(PasteType.ALL, PasteOperation.NONE, false, true);
	}
	
	protected boolean doCut() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		
		setClipboard(new Clipboard(Clipboard.Type.CUT, sheet.getSheetName(), selection));
		return true;
	}
	
	protected boolean doFontFamily(String fontFamily) {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		CellOperationUtil.applyFontName(range, fontFamily);
		return true;
	}
	
	protected boolean doFontSize(int fontSize) {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		CellOperationUtil.applyFontSize(range, (short)fontSize);
		return true;
	}
	
	protected boolean doFontBold() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}

		//toggle and apply bold of first cell to dest
		Boldweight bw = range.getCellStyle().getFont().getBoldweight();
		if(Boldweight.BOLD.equals(bw)){
			bw = Boldweight.NORMAL;
		}else{
			bw = Boldweight.BOLD;
		}
		
		CellOperationUtil.applyFontBoldweight(range, bw);
		return true;
	}
	
	protected boolean doFontItalic() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}

		//toggle and apply bold of first cell to dest
		boolean italic = !range.getCellStyle().getFont().isItalic();
		CellOperationUtil.applyFontItalic(range, italic);
		return true;
	}
	
	protected boolean doFontStrikeout() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}

		//toggle and apply bold of first cell to dest
		boolean strikeout = !range.getCellStyle().getFont().isStrikeout();
		CellOperationUtil.applyFontStrikeout(range, strikeout);
		return true;
	}
	
	protected boolean doFontUnderline() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}

		//toggle and apply bold of first cell to dest
		Underline underline = range.getCellStyle().getFont().getUnderline();
		if(Underline.NONE.equals(underline)){
			underline = Underline.SINGLE;
		}else{
			underline = Underline.NONE;
		}
		
		CellOperationUtil.applyFontUnderline(range, underline);	
		return true;
	}

	protected boolean doBorder(ApplyBorderType type,BorderType borderTYpe, String color){
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		CellOperationUtil.applyBorder(range,type, borderTYpe, color);
		return true;
	}
	
	protected boolean doBorder(String color) {
		return doBorder(ApplyBorderType.EDGE_BOTTOM,BorderType.MEDIUM,color);
	}
	
	protected boolean doBorderBottom(String color) {
		return doBorder(ApplyBorderType.EDGE_BOTTOM,BorderType.MEDIUM,color);
	}
	
	protected boolean doBoderTop(String color) {
		return doBorder(ApplyBorderType.EDGE_TOP,BorderType.MEDIUM,color);
	}
	
	protected boolean doBorderLeft(String color) {
		return doBorder(ApplyBorderType.EDGE_LEFT,BorderType.MEDIUM,color);
	}
	
	protected boolean doBorderRight(String color) {
		return doBorder(ApplyBorderType.EDGE_RIGHT,BorderType.MEDIUM,color);
	}
	
	protected boolean doBorderNo(String color) {
		return doBorder(ApplyBorderType.FULL,BorderType.NONE,color);
	}
	
	protected boolean doBorderAll(String color) {
		return doBorder(ApplyBorderType.FULL,BorderType.MEDIUM,color);
	}
	
	protected boolean doBorderOutside(String color) {
		return doBorder(ApplyBorderType.OUTLINE,BorderType.MEDIUM,color);
	}
	
	protected boolean doBorderInside(String color) {
		return doBorder(ApplyBorderType.INSIDE,BorderType.MEDIUM,color);
	}
	
	protected boolean doBorderInsideHorizontal(String color) {
		return doBorder(ApplyBorderType.INSIDE_HORIZONTAL,BorderType.MEDIUM,color);
	}
	
	protected boolean doBorderInsideVertical(String color) {
		return doBorder(ApplyBorderType.INSIDE_VERTICAL,BorderType.MEDIUM,color);
	}
	
	protected boolean doFontColor(String color) {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		CellOperationUtil.applyFontColor(range, color);
		return true;
	}
	
	protected boolean doFillColor(String color) {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		CellOperationUtil.applyBackgroundColor(range,color);
		return true;
	}
	
	protected boolean doVerticalAlignTop() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		CellOperationUtil.applyVerticalAlignment(range, VerticalAlignment.TOP);
		return true;
	}
	
	protected boolean doVerticalAlignMiddle() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		CellOperationUtil.applyVerticalAlignment(range, VerticalAlignment.CENTER);
		return true;
	}

	protected boolean doVerticalAlignBottom() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		CellOperationUtil.applyVerticalAlignment(range, VerticalAlignment.BOTTOM);
		return true;
	}
	
	protected boolean doHorizontalAlignLeft() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		CellOperationUtil.applyAlignment(range, Alignment.LEFT);
		return true;
	}
	
	protected boolean doHorizontalAlignCenter() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		CellOperationUtil.applyAlignment(range, Alignment.CENTER);
		return true;
	}
	
	protected boolean doHorizontalAlignRight() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		CellOperationUtil.applyAlignment(range, Alignment.RIGHT);
		return true;
	}
	
	protected boolean doWrapText() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		boolean wrapped = !range.getCellStyle().isWrapText();
		CellOperationUtil.applyWrapText(range, wrapped);
		return true;
	}
	
	protected boolean doMergeAndCenter() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		CellOperationUtil.toggleMergeCenter(range);
		clearClipboard();
		return true;
	}
	
	protected boolean doMergeAcross() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		CellOperationUtil.merge(range, true);
		clearClipboard();
		return true;
	}
	
	protected boolean doMergeCell() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		CellOperationUtil.merge(range, false);
		clearClipboard();
		return true;
	}
	
	protected boolean doUnmergeCell() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		CellOperationUtil.unMerge(range);
		clearClipboard();
		return true;
	}
	
	protected boolean doShiftCellRight() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		CellOperationUtil.insert(range,InsertShift.RIGHT, InsertCopyOrigin.FORMAT_RIGHT_BELOW);
		clearClipboard();
		return true;
	}
	
	protected boolean doShiftCellDown() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		CellOperationUtil.insert(range,InsertShift.DOWN, InsertCopyOrigin.FORMAT_LEFT_ABOVE);
		clearClipboard();
		return true;
	}
	
	protected boolean doInsertSheetRow() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		
		if(range.isWholeColumn()){
			showWarnMessage("don't allow to inser row when select whole column");
			return true;
		}
		
		range = range.toRowRange();
		CellOperationUtil.insert(range,InsertShift.DOWN, InsertCopyOrigin.FORMAT_LEFT_ABOVE);
		clearClipboard();
		return true;
	}
	
	protected boolean doInsertSheetColumn() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		
		if(range.isWholeRow()){
			showWarnMessage("don't allow to inser column when select whole row");
			return true;
		}
		
		range = range.toColumnRange();
		CellOperationUtil.insert(range,InsertShift.RIGHT, InsertCopyOrigin.FORMAT_RIGHT_BELOW);
		clearClipboard();
		return true;
	}
	
	protected boolean doShiftCellLeft() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		
		CellOperationUtil.delete(range,DeleteShift.LEFT);
		clearClipboard();
		return true;
	}
	
	protected boolean doShiftCellUp() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		
		CellOperationUtil.delete(range,DeleteShift.UP);
		clearClipboard();
		return true;
	}
	
	protected boolean doDeleteSheetRow() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		
		if(range.isWholeColumn()){
			showWarnMessage("don't allow to delete all rows");
			return true;
		}
		
		range = range.toRowRange();
		CellOperationUtil.delete(range, DeleteShift.UP);
		clearClipboard();
		return true;
	}
	
	protected boolean doDeleteSheetColumn() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		
		if(range.isWholeRow()){
			showWarnMessage("don't allow to delete all column");
			return true;
		}
		
		range = range.toColumnRange();
		CellOperationUtil.delete(range, DeleteShift.LEFT);
		clearClipboard();
		return true;
	}

	protected boolean doClearStyle() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		CellOperationUtil.clearStyles(range);
		return true;
	}
	
	protected boolean doClearContent() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		CellOperationUtil.clearContents(range);
		return true;
	}
	
	protected boolean doClearAll() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		CellOperationUtil.clearAll(range);
		return true;
	}
	
	protected boolean doSortAscending() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		CellOperationUtil.sort(range,false);
		clearClipboard();
		return true;
	}
	
	protected boolean doSortDescending() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		CellOperationUtil.sort(range,true);
		clearClipboard();
		return true;
	}
	
	protected boolean doFilter() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		SheetOperationUtil.toggleAutoFilter(range);
		clearClipboard();
		return true;
	}
	
	protected boolean doClearFilter() {
		Sheet sheet = getSheet();
		
		Range range = Ranges.range(sheet);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		SheetOperationUtil.resetAutoFilter(range);
		clearClipboard();
		return true;
	}
	
	protected boolean doReapplyFilter() {
		Sheet sheet = getSheet();
		
		Range range = Ranges.range(sheet);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		SheetOperationUtil.applyAutoFilter(range);
		clearClipboard();
		return true;
	}
	
	protected boolean doProtectSheet() {
		
		Sheet sheet = getSheet();
		
		Range range = Ranges.range(sheet);
		
		String newpassword = "1234";//TODO, make it meaningful
		if(range.isProtected()){
			SheetOperationUtil.protectSheet(range,null,null);
		}else{
			SheetOperationUtil.protectSheet(range,null,newpassword);
		}
		return true;
	}
	
	protected boolean doGridlines() {
		Sheet sheet = getSheet();
		
		Range range = Ranges.range(sheet);
		
		SheetOperationUtil.displaySheetGridlines(range,!range.isDisplaySheetGridlines());
		return true;
	}

	private static <T> T checkNotNull(String message, T t) {
		if (t == null) {
			throw new NullPointerException(message);
		}
		return t;
	}

	public static class UserActionContext {

		final Spreadsheet spreadsheet;
		final Sheet sheet;
		final String action;
		final Rect selection;
		final Map extraData;
		
		public UserActionContext(Spreadsheet spreadsheet, Sheet sheet,String action,
				Rect selection, Map extraData) {
			this.spreadsheet = spreadsheet;
			this.sheet = sheet;
			this.action = action;
			this.selection = selection;
			this.extraData = extraData;
		}
	}
	
	/**
	 * Clipboard data object for copy/paste
	 * @author dennis
	 *
	 */
	public static class Clipboard {
		public enum Type {
			COPY,
			CUT
		}
		
		public final Type type;
		public final Rect sourceRect;
		public final String sourceSheetName;
		
		public Clipboard(Type type, String sourceSheetName,Rect sourceRect) {
			this.type = checkNotNull("Clipboard's type cannot be null", type);
			this.sourceSheetName = checkNotNull("Clipboard's sourceSheetName cannot be null", sourceSheetName);
			this.sourceRect = checkNotNull("Clipboard's sourceRect cannot be null", sourceRect);
		}
	}
	
	// non-implemented action
	
//	protected boolean doRowHeight() {
//		showNotImplement(UserAction.ROW_HEIGHT.getAction());
//		return true;
//	}
//
//	protected boolean doColumnWidth() {
//		showNotImplement(UserAction.COLUMN_WIDTH.getAction());
//		return true;
//	}
//
//	protected boolean doFormatCell() {
//		showNotImplement(UserAction.FORMAT_CELL.getAction());
//		return true;
//	}
//
//	protected boolean doHyperlink() {
//		showNotImplement(UserAction.HYPERLINK.getAction());
//		return true;
//	}
//
//	protected boolean doCustomSort() {
//		showNotImplement(UserAction.CUSTOM_SORT.getAction());
//		return true;
//	}
//
//	protected boolean doPasteSpecial() {
//		showNotImplement(UserAction.PASTE_SPECIAL.getAction());
//		return true;
//	}
//
//	protected boolean doExportPDF() {
//		showNotImplement(UserAction.EXPORT_PDF.getAction());
//		return true;
//	}
//
//	protected boolean doSaveBook() {
//		showNotImplement(UserAction.SAVE_BOOK.getAction());
//		return true;
//	}
//
//	protected boolean doNewBook() {
//		showNotImplement(UserAction.NEW_BOOK.getAction());
//		return true;
//	}
//
//	protected boolean doShowFormulaPanel() {
//		showNotImplement(UserAction.FORMULA_PANEL.getAction());
//		return true;
//	}
//
//	protected boolean doShowInsertPanel() {
//		showNotImplement(UserAction.INSERT_PANEL.getAction());
//		return true;
//	}
//
//	protected boolean doShowHomePanel() {
//		showNotImplement(UserAction.HOME_PANEL.getAction());
//		return true;
//	}
//
//	protected boolean doInsertFunction() {
//		showNotImplement(UserAction.INSERT_FUNCTION.getAction());
//		return true;
//	}

	@Override
	public Set<String> getInterestedEvents() {
		Set<String> evts = new LinkedHashSet<String>();
		evts.add(Events.ON_AUX_ACTION);
		evts.add(Events.ON_SHEET_SELECT);
		evts.add(Events.ON_CTRL_KEY);
		evts.add(Events.ON_CELL_SELECTION_UPDATE);
		evts.add(org.zkoss.zk.ui.event.Events.ON_CANCEL);
//		evts.add(Events.ON_CELL_DOUBLE_CLICK);
		evts.add(Events.ON_START_EDITING);
		return evts;
	}
	
	@Override
	public void onEvent(Event event) throws Exception {
		Component comp = event.getTarget();
		if(!(comp instanceof Spreadsheet)) return;
		
		Spreadsheet spreadsheet = (Spreadsheet)comp;
		Sheet sheet = spreadsheet.getSelectedSheet();
		String action = "";
		Rect selection;
		Map extraData = null;
		if(event instanceof KeyEvent){
			//respect zss key-even't selection 
			//(could consider to remove this extra spec some day)
			selection = ((KeyEvent)event).getSelection();
//			action = "keyStroke";
		}else if(event instanceof AuxActionEvent){
			AuxActionEvent evt = ((AuxActionEvent)event);
			selection = evt.getSelection();
			//use the sheet form aux, it could be not the selected sheet
			sheet = evt.getSheet();
			action = evt.getAction();
			extraData = evt.getExtraData();
		}else if(event instanceof CellSelectionUpdateEvent){
			CellSelectionUpdateEvent evt = ((CellSelectionUpdateEvent)event);
			selection = new Rect(evt.getLeft(),evt.getTop(),evt.getRight(),evt.getBottom());
		}else{
			selection = spreadsheet.getSelection();
		}
		
		Rect visibleSelection = new Rect(selection.getLeft(), selection.getTop(), Math.min(
			    spreadsheet.getMaxVisibleColumns(), selection.getRight()), Math.min(
			    spreadsheet.getMaxVisibleRows(), selection.getBottom()));
		
		final UserActionContext ctx = new UserActionContext(spreadsheet, sheet,action,visibleSelection,extraData);
		setContext(ctx);
		try{
			if(event instanceof AuxActionEvent){
				dispatchAction(action);
			}else{
				onEventAnother(event);
			}
		}finally{
			releaseContext();
		}
	}
	
	protected void setContext(UserActionContext ctx){
		_ctx.set(ctx);
	}
	protected UserActionContext getContext(){
		return _ctx.get();
	}
	protected void releaseContext(){
		_ctx.set(null);
	}

	private void onEventAnother(Event event) throws Exception {

		String nm = event.getName();
		if(Events.ON_SHEET_SELECT.equals(nm)){
			
			updateClipboardEffect(getSheet());
			//TODO 20130513, Dennis, looks like I don't need to do this here?
			//syncAutoFilter();
			
			//TODO this should be spreadsheet's job
			//toggleActionOnSheetSelected() 
		}else if(Events.ON_CTRL_KEY.equals(nm)){
			KeyEvent kevt = (KeyEvent)event;
			boolean r = doKeystroke(kevt.getKeyCode(), kevt.isCtrlKey(), kevt.isShiftKey(), kevt.isAltKey());
			if(r){
				//to disable client copy/paste feature if there is any server side copy/paste
				if(kevt.isCtrlKey() && kevt.getKeyCode()=='V'){
					//internal only
					getSpreadsheet().smartUpdate("doPasteFromServer", true);
				}
			}
		}else if(Events.ON_CELL_SELECTION_UPDATE.equals(nm)){
			CellSelectionUpdateEvent evt = (CellSelectionUpdateEvent)event;
			//last selection either get form selection or from event
			if(evt.getAction()==CellSelectionAction.MOVE){
				doMoveCellSelection(new Rect(evt.getOrigleft(),evt.getOrigtop(),evt.getOrigright(),evt.getOrigbottom()),getSelection());
			}else if(evt.getAction()==CellSelectionAction.RESIZE){
				doResizeCellSelection(new Rect(evt.getOrigleft(),evt.getOrigtop(),evt.getOrigright(),evt.getOrigbottom()),getSelection());
			}
		/*}else if(Events.ON_CELL_DOUBLE_CLICK.equals(nm)){//TODO check if we need it still
			clearClipboard();
		*/}else if(Events.ON_START_EDITING.equals(nm)){
			clearClipboard();
		}else if(org.zkoss.zk.ui.event.Events.ON_CANCEL.equals(nm)){
			clearClipboard();
		}
	}
	
	protected boolean doMoveCellSelection(Rect original,Rect selection) {
		Sheet sheet = getSheet();
		Range src = Ranges.range(sheet,original.getTop(),original.getLeft(),original.getBottom(),original.getRight());
		Range dest = Ranges.range(sheet,selection.getTop(),selection.getLeft(),selection.getBottom(),selection.getRight());
		
		if(dest.isProtected()){
			showProtectMessage();
			return true;
		}
		
		final int nRow = selection.getTop() - original.getTop();
		final int nCol = selection.getLeft() - original.getLeft();
		CellOperationUtil.shift(src,nRow, nCol);
		
		return true;
	}
	
	protected boolean doResizeCellSelection(Rect original,Rect selection) {
		Sheet sheet = getSheet();
		Range src = Ranges.range(sheet,original.getTop(),original.getLeft(),original.getBottom(),original.getRight());
		Range dest = Ranges.range(sheet,selection.getTop(),selection.getLeft(),selection.getBottom(),selection.getRight());
		
		if(dest.isProtected()){
			showProtectMessage();
			return true;
		}
		
		SheetOperationUtil.autoFill(src,dest, AutoFillType.DEFAULT);	
		
		return true;
	}

	protected void updateClipboardEffect(Sheet sheet) {
		//to sync the 
		Clipboard cb = getClipboard();
		if (cb != null) {
			//TODO a way to know the book is different already?
			final Book book = sheet.getBook();
			final Sheet src = book.getSheet(cb.sourceSheetName);
			if(src==null){
				clearClipboard();
			}else{
				if(sheet.equals(src)){
					getSpreadsheet().setHighlight(cb.sourceRect);
				}else{
					getSpreadsheet().setHighlight(null);
				}
			}
		}
	}

	@Override
	public String getCtrlKeys() {
		return "^X^C^V^B^I^U#del";
	}

	@Override
	public Set<String> getSupportedUserAction(Sheet sheet) {
		Set<String> actions = new LinkedHashSet<String>();
		if(sheet==null){
			//
			return actions;
//			disabled.addAll(Arrays.asList(DefaultUserActionHandler.DISABLED_ACTION_WHEN_BOOK_CLOSE));
		}else{
			int sheetnum = sheet.getBook().getNumberOfSheets();
//			actions.add(UserAction.HOME_PANEL.getAction());
//			actions.add(UserAction.INSERT_PANEL.getAction());
//			actions.add(UserAction.FORMULA_PANEL.getAction());
//			actions.add(UserAction.NEW_BOOK.getAction());
//			actions.add(UserAction.SAVE_BOOK.getAction());
//			actions.add(UserAction.EXPORT_PDF.getAction());
			
			actions.add(DefaultUserAction.CLOSE_BOOK.getAction());
			
			actions.add(DefaultUserAction.ADD_SHEET.getAction());
			if(sheetnum>1){
				actions.add(DefaultUserAction.DELETE_SHEET.getAction());
				if(sheet.getBook().getSheetIndex(sheet)>0){
					actions.add(DefaultUserAction.MOVE_SHEET_LEFT.getAction());
				}
				if(sheet.getBook().getSheetIndex(sheet)<sheetnum-1){
					actions.add(DefaultUserAction.MOVE_SHEET_RIGHT.getAction());
				}
				
			}
			actions.add(DefaultUserAction.RENAME_SHEET.getAction());
			actions.add(DefaultUserAction.PROTECT_SHEET.getAction());
			actions.add(DefaultUserAction.GRIDLINES.getAction());

			
			
			if(!sheet.isProtected()){
				actions.add(DefaultUserAction.PASTE.getAction());
				actions.add(DefaultUserAction.PASTE_FORMULA.getAction());
				actions.add(DefaultUserAction.PASTE_VALUE.getAction());
				actions.add(DefaultUserAction.PASTE_ALL_EXPECT_BORDERS.getAction());
				actions.add(DefaultUserAction.PASTE_TRANSPOSE.getAction());
//				actions.add(UserAction.PASTE_SPECIAL.getAction());
				actions.add(DefaultUserAction.CUT.getAction());
				actions.add(DefaultUserAction.COPY.getAction());
				actions.add(DefaultUserAction.FONT_FAMILY.getAction());
				actions.add(DefaultUserAction.FONT_SIZE.getAction());
				actions.add(DefaultUserAction.FONT_BOLD.getAction());
				actions.add(DefaultUserAction.FONT_ITALIC.getAction());
				actions.add(DefaultUserAction.FONT_UNDERLINE.getAction());
				actions.add(DefaultUserAction.FONT_STRIKE.getAction());
				actions.add(DefaultUserAction.BORDER.getAction());
				actions.add(DefaultUserAction.BORDER_BOTTOM.getAction());
				actions.add(DefaultUserAction.BORDER_TOP.getAction());
				actions.add(DefaultUserAction.BORDER_LEFT.getAction());
				actions.add(DefaultUserAction.BORDER_RIGHT.getAction());
				actions.add(DefaultUserAction.BORDER_NO.getAction());
				actions.add(DefaultUserAction.BORDER_ALL.getAction());
				actions.add(DefaultUserAction.BORDER_OUTSIDE.getAction());
				actions.add(DefaultUserAction.BORDER_INSIDE.getAction());
				actions.add(DefaultUserAction.BORDER_INSIDE_HORIZONTAL.getAction());
				actions.add(DefaultUserAction.BORDER_INSIDE_VERTICAL.getAction());
				actions.add(DefaultUserAction.FONT_COLOR.getAction());
				actions.add(DefaultUserAction.FILL_COLOR.getAction());
				actions.add(DefaultUserAction.VERTICAL_ALIGN.getAction());//folder
				actions.add(DefaultUserAction.VERTICAL_ALIGN_TOP.getAction());
				actions.add(DefaultUserAction.VERTICAL_ALIGN_MIDDLE.getAction());
				actions.add(DefaultUserAction.VERTICAL_ALIGN_BOTTOM.getAction());
				actions.add(DefaultUserAction.HORIZONTAL_ALIGN.getAction());//folder
				actions.add(DefaultUserAction.HORIZONTAL_ALIGN_LEFT.getAction());
				actions.add(DefaultUserAction.HORIZONTAL_ALIGN_CENTER.getAction());
				actions.add(DefaultUserAction.HORIZONTAL_ALIGN_RIGHT.getAction());
				actions.add(DefaultUserAction.WRAP_TEXT.getAction());
				actions.add(DefaultUserAction.MERGE_AND_CENTER.getAction());
				actions.add(DefaultUserAction.MERGE_ACROSS.getAction());
				actions.add(DefaultUserAction.MERGE_CELL.getAction());
				actions.add(DefaultUserAction.UNMERGE_CELL.getAction());
				actions.add(DefaultUserAction.INSERT.getAction());//folder
				actions.add(DefaultUserAction.INSERT_SHIFT_CELL_RIGHT.getAction());
				actions.add(DefaultUserAction.INSERT_SHIFT_CELL_DOWN.getAction());
				actions.add(DefaultUserAction.INSERT_SHEET_ROW.getAction());
				actions.add(DefaultUserAction.INSERT_SHEET_COLUMN.getAction());
				actions.add(DefaultUserAction.DELETE.getAction());//folder
				actions.add(DefaultUserAction.DELETE_SHIFT_CELL_LEFT.getAction());
				actions.add(DefaultUserAction.DELETE_SHIFT_CELL_UP.getAction());
				actions.add(DefaultUserAction.DELETE_SHEET_ROW.getAction());
				actions.add(DefaultUserAction.DELETE_SHEET_COLUMN.getAction());
				actions.add(DefaultUserAction.SORT_AND_FILTER.getAction());//folder
				actions.add(DefaultUserAction.SORT_ASCENDING.getAction());
				actions.add(DefaultUserAction.SORT_DESCENDING.getAction());
//				actions.add(UserAction.CUSTOM_SORT.getAction());
				actions.add(DefaultUserAction.CLEAR.getAction());//folder
				actions.add(DefaultUserAction.CLEAR_CONTENT.getAction());
				actions.add(DefaultUserAction.CLEAR_STYLE.getAction());
				actions.add(DefaultUserAction.CLEAR_ALL.getAction());
//				actions.add(UserAction.HYPERLINK.getAction());
//				actions.add(UserAction.FORMAT_CELL.getAction());
//				actions.add(UserAction.COLUMN_WIDTH.getAction());
//				actions.add(UserAction.ROW_HEIGHT.getAction());
				actions.add(DefaultUserAction.HIDE_COLUMN.getAction());
				actions.add(DefaultUserAction.UNHIDE_COLUMN.getAction());
				actions.add(DefaultUserAction.HIDE_ROW.getAction());
				actions.add(DefaultUserAction.UNHIDE_ROW.getAction());
//				actions.add(UserAction.INSERT_FUNCTION.getAction());				
			}
			
			actions.add(DefaultUserAction.FILTER.getAction());
			if(sheet.isAutoFilterEnabled()){
				actions.add(DefaultUserAction.CLEAR_FILTER.getAction());
				actions.add(DefaultUserAction.REAPPLY_FILTER.getAction());
//				disabled.addAll(Arrays.asList(DefaultUserActionHandler.DISABLED_ACTION_WHEN_FILTER_OFF));
			}
		}
		return actions;
	}

}
