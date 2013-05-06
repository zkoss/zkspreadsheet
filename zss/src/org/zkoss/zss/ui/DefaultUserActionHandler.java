package org.zkoss.zss.ui;

import java.util.Map;

import org.zkoss.image.AImage;
import org.zkoss.lang.Strings;
import org.zkoss.util.media.Media;
import org.zkoss.util.resource.Labels;
import org.zkoss.zk.ui.Desktop;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zk.ui.event.UploadEvent;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.SheetAnchor;
import org.zkoss.zss.api.SheetOperationUtil;
import org.zkoss.zss.api.Range.ApplyBorderType;
import org.zkoss.zss.api.Range.DeleteShift;
import org.zkoss.zss.api.Range.InsertCopyOrigin;
import org.zkoss.zss.api.Range.InsertShift;
import org.zkoss.zss.api.Range.PasteOperation;
import org.zkoss.zss.api.Range.PasteType;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Chart;
import org.zkoss.zss.api.model.ChartData;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.CellStyle.Alignment;
import org.zkoss.zss.api.model.CellStyle.BorderType;
import org.zkoss.zss.api.model.CellStyle.VerticalAlignment;
import org.zkoss.zss.api.model.Font.Boldweight;
import org.zkoss.zss.api.model.Font.Underline;
import org.zkoss.zss.ui.DefaultUserActionHandler.Clipboard.Type;
import org.zkoss.zss.ui.event.KeyEvent;
import org.zkoss.zul.Fileupload;
import org.zkoss.zul.Messagebox;

public class DefaultUserActionHandler implements UserActionHandler {
	
	private static ThreadLocal<UserActionContext> _ctx = new ThreadLocal<UserActionContext>();
	private Clipboard _clipboard;

	
	protected Spreadsheet getSpreadsheet(){
		return _ctx.get().spreadsheet;
	}
	
	protected Sheet getSheet(){
		return _ctx.get().sheet;
	}
	
	protected Book getBook(){
		return getSheet().getBook();
	}
	
	protected Rect getSelection(){
		return _ctx.get().selection;
	}
	
	@Override
	public void handleAction(Spreadsheet spreadsheet, Sheet targetSheet,
			String action, Rect selection, Map<String, Object> extraData) {
		final UserActionContext ctx = new UserActionContext(spreadsheet, targetSheet,action,selection,extraData);
		_ctx.set(ctx);
		try{
			dispatchAction(spreadsheet,targetSheet,action,selection,extraData);
		}finally{
			_ctx.set(null);
		}
	}
	
	
	protected void dispatchAction(Spreadsheet spreadsheet, Sheet targetSheet,
			String action, Rect selection, Map<String, Object> extraData) {
		
		if (Action.HOME_PANEL.toString().equals(action)) {
			doShowHomePanel();
		} else if (Action.INSERT_PANEL.equals(action)) {
			doShowInsertPanel();
		} else if (Action.FORMULA_PANEL.equals(action)) {
			doShowFormulaPanel();
		} else if (Action.ADD_SHEET.toString().equals(action)) {
			doAddSheet();
		} else if (Action.DELETE_SHEET.equals(action)) {
			doDeleteSheet();
		} else if (Action.RENAME_SHEET.equals(action)) {
			String name = (String) extraData.get("name");
			doRenameSheet(name);
		} else if (Action.MOVE_SHEET_LEFT.equals(action)) {
			doMoveSheetLeft();
		} else if (Action.MOVE_SHEET_RIGHT.equals(action)) {
			doMoveSheetRight();
		} else if (Action.PROTECT_SHEET.equals(action)) {
			doProtectSheet();
		} else if (Action.GRIDLINES.equals(action)) {
			doGridlines();
		} else if (Action.NEW_BOOK.equals(action)) {
			doNewBook();
		} else if (Action.SAVE_BOOK.equals(action)) {
			doSaveBook();
		} else if (Action.EXPORT_PDF.equals(action)) {
			doExportPDF();
		} else if (Action.PASTE.equals(action)) {
			doPaste();
		} else if (Action.PASTE_FORMULA.equals(action)) {
			doPasteFormula();
		} else if (Action.PASTE_VALUE.equals(action)) {
			doPasteValue();
		} else if (Action.PASTE_ALL_EXPECT_BORDERS.equals(action)) {
			doPasteAllExceptBorder();
		} else if (Action.PASTE_TRANSPOSE.equals(action)) {
			doPasteTranspose();
		} else if (Action.PASTE_SPECIAL.equals(action)) {
			doPasteSpecial();
		} else if (Action.CUT.equals(action)) {
			doCut();	
		} else if (Action.COPY.equals(action)) {
			doCopy();
		} else if (Action.FONT_FAMILY.equals(action)) {
			doFontFamily((String)extraData.get("name"));
		} else if (Action.FONT_SIZE.equals(action)) {
			Integer fontSize = Integer.parseInt((String)extraData.get("size"));
			doFontSize(fontSize);
		} else if (Action.FONT_BOLD.equals(action)) {
			doFontBold();
		} else if (Action.FONT_ITALIC.equals(action)) {
			doFontItalic();
		} else if (Action.FONT_UNDERLINE.equals(action)) {
			doFontUnderline();
		} else if (Action.FONT_STRIKE.equals(action)) {
			doFontStrikeout();
		} else if (Action.BORDER.equals(action)) {
			doBorder(getBorderColor(extraData));
		} else if (Action.BORDER_BOTTOM.equals(action)) {
			doBorderBottom(getBorderColor(extraData));
		} else if (Action.BORDER_TOP.equals(action)) {
			doBoderTop(getBorderColor(extraData));
		} else if (Action.BORDER_LEFT.equals(action)) {
			doBorderLeft(getBorderColor(extraData));
		} else if (Action.BORDER_RIGHT.equals(action)) {
			doBorderRight(getBorderColor(extraData));
		} else if (Action.BORDER_NO.equals(action)) {
			doBorderNo(getBorderColor(extraData));
		} else if (Action.BORDER_ALL.equals(action)) {
			doBorderAll(getBorderColor(extraData));
		} else if (Action.BORDER_OUTSIDE.equals(action)) {
			doBorderOutside(getBorderColor(extraData));
		} else if (Action.BORDER_INSIDE.equals(action)) {
			doBorderInside(getBorderColor(extraData));
		} else if (Action.BORDER_INSIDE_HORIZONTAL.equals(action)) {
			doBorderInsideHorizontal(getBorderColor(extraData));
		} else if (Action.BORDER_INSIDE_VERTICAL.equals(action)) {
			doBorderInsideVertical(getBorderColor(extraData));
		} else if (Action.FONT_COLOR.equals(action)) {
			doFontColor(getFontColor(extraData));
		} else if (Action.FILL_COLOR.equals(action)) {
			doFillColor(getFillColor(extraData));
		} else if (Action.VERTICAL_ALIGN_TOP.equals(action)) {
			doVerticalAlignTop();
		} else if (Action.VERTICAL_ALIGN_MIDDLE.equals(action)) {
			doVerticalAlignMiddle();
		} else if (Action.VERTICAL_ALIGN_BOTTOM.equals(action)) {
			doVerticalAlignBottom();
		} else if (Action.HORIZONTAL_ALIGN_LEFT.equals(action)) {
			doHorizontalAlignLeft();
		} else if (Action.HORIZONTAL_ALIGN_CENTER.equals(action)) {
			doHorizontalAlignCenter();
		} else if (Action.HORIZONTAL_ALIGN_RIGHT.equals(action)) {
			doHorizontalAlignRight();
		} else if (Action.WRAP_TEXT.equals(action)) {
			doWrapText();
		} else if (Action.MERGE_AND_CENTER.equals(action)) {
			doMergeAndCenter();
		} else if (Action.MERGE_ACROSS.equals(action)) {
			doMergeAcross();
		} else if (Action.MERGE_CELL.equals(action)) {
			doMergeCell();
		} else if (Action.UNMERGE_CELL.equals(action)) {
			doUnmergeCell();
		} else if (Action.INSERT_SHIFT_CELL_RIGHT.equals(action)) {
			doShiftCellRight();
		} else if (Action.INSERT_SHIFT_CELL_DOWN.equals(action)) {
			doShiftCellDown();
		} else if (Action.INSERT_SHEET_ROW.equals(action)) {
			doInsertSheetRow();
		} else if (Action.INSERT_SHEET_COLUMN.equals(action)) {
			doInsertSheetColumn();
		} else if (Action.DELETE_SHIFT_CELL_LEFT.equals(action)) {
			doShiftCellLeft();
		} else if (Action.DELETE_SHIFT_CELL_UP.equals(action)) {
			doShiftCellUp();
		} else if (Action.DELETE_SHEET_ROW.equals(action)) {
			doDeleteSheetRow();
		} else if (Action.DELETE_SHEET_COLUMN.equals(action)) {
			doDeleteSheetColumn();
		} else if (Action.SORT_ASCENDING.equals(action)) {
			doSortAscending();
		} else if (Action.SORT_DESCENDING.equals(action)) {
			doSortDescending();
		} else if (Action.CUSTOM_SORT.equals(action)) {
			doCustomSort();
		} else if (Action.FILTER.equals(action)) {
			doFilter();
		} else if (Action.CLEAR_FILTER.equals(action)) {
			doClearFilter();
		} else if (Action.REAPPLY_FILTER.equals(action)) {
			doReapplyFilter();
		} else if (Action.CLEAR_CONTENT.equals(action)) {
			doClearContent();
		} else if (Action.CLEAR_STYLE.equals(action)) {
			doClearStyle();
		} else if (Action.CLEAR_ALL.equals(action)) {
			doClearAll();
		} else if (Action.COLUMN_CHART.equals(action)) {
			doColumnChart();
		} else if (Action.COLUMN_CHART_3D.equals(action)) {
			doColumnChart3D();
		} else if (Action.LINE_CHART.equals(action)) {
			doLineChart();
		} else if (Action.LINE_CHART_3D.equals(action)) {
			doLineChart3D();
		} else if (Action.PIE_CHART.equals(action)) {
			doPieChart();
		} else if (Action.PIE_CHART_3D.equals(action)) {
			doPieChart3D();
		} else if (Action.BAR_CHART.equals(action)) {
			doBarChart();
		} else if (Action.BAR_CHART_3D.equals(action)) {
			doBarChart3D();
		} else if (Action.AREA_CHART.equals(action)) {
			doAreaChart();
		} else if (Action.SCATTER_CHART.equals(action)) {
			doScatterChart();
		} else if (Action.DOUGHNUT_CHART.equals(action)) {
			doDoughnutChart();
		} else if (Action.HYPERLINK.equals(action)) {
			doHyperlink();
		} else if (Action.INSERT_PICTURE.equals(action)) {
			doInsertPicture();
		} else if (Action.CLOSE_BOOK.equals(action)) {
			doCloseBook();
		} else if (Action.FORMAT_CELL.equals(action)) {
			doFormatCell();
		} else if (Action.COLUMN_WIDTH.equals(action)) {
			doColumnWidth();
		} else if (Action.ROW_HEIGHT.equals(action)) {
			doRowHeight();
		} else if (Action.HIDE_COLUMN.equals(action)) {
			doHideColumn();
		} else if (Action.UNHIDE_COLUMN.equals(action)) {
			doUnhideColumn();
		} else if (Action.HIDE_ROW.equals(action)) {
			doHideRow();
		} else if (Action.UNHIDE_ROW.equals(action)) {
			doUnhideRow();
		} else if (Action.INSERT_FUNCTION.equals(action)) {
			doInsertFunction();
		}else{
			showNotImplement(action);
		}
	}

	protected void doMoveSheetRight() {
		Book book = getBook();
		Sheet sheet = getSheet();
		
		int max = book.getNumberOfSheets();
		int i = book.getSheetIndex(sheet);
		
		if(i<max){
			i ++;
		}else{
			//do nothing
			return;
		}
		
		Range range = Ranges.range(sheet);
		SheetOperationUtil.setSheetOrder(range,i);
	}

	protected void doMoveSheetLeft() {
		Book book = getBook();
		Sheet sheet = getSheet();
		
		if(book.getSheetIndex(sheet)==0){
			//do nothing
			return;
		}
		
		Range range = Ranges.range(sheet);
		SheetOperationUtil.setSheetOrder(range,0);
		
		
	}

	protected void doRenameSheet(String newname) {
		Book book = getBook();
		Sheet sheet = getSheet();

		if(book.getSheet(newname)!=null){
			showWarnMessage("Canot rename a sheet to the same as another.");
			return;
		}
		
		Range range = Ranges.range(sheet);
		SheetOperationUtil.renameSheet(range,newname);
	}

	protected void doDeleteSheet() {
		Book book = getBook();
		Sheet sheet = getSheet();
		
		int num = book.getNumberOfSheets();
		if(num<=1){
			showWarnMessage("Canot delete last sheet.");
			return;
		}
		int index = book.getSheetIndex(sheet);
		
		Range range = Ranges.range(sheet);
		SheetOperationUtil.deleteSheet(range);
		
		if(index==num-1){
			index--;
		}
		
		getSpreadsheet().setSelectedSheet(book.getSheetAt(index).getSheetName());
	}

	protected void doAddSheet() {
		String prefix = Labels.getLabel(Action.SHEET.getLabelKey());
		if (Strings.isEmpty(prefix))
			prefix = "Sheet";
		Sheet sheet = getSheet();

		Range range = Ranges.range(sheet);
		SheetOperationUtil.addSheet(range,prefix);
	}

	/**
	 * Returns the border color
	 * @return
	 */
	protected String getDefaultBorderColor() {
		return "#000000";
	}
	
	
	/**
	 * Returns the border color
	 * @return
	 */
	protected String getDefaultFontColor() {
		return "#000000";
	}
	
	
	/**
	 * Returns the border color
	 * @return
	 */
	protected String getDefaultFillColor() {
		return "#FFFFFF";
	}
	
	private String getBorderColor(Map extraData){
		String color = (String)extraData.get("color");
		if (Strings.isEmpty(color)) {//CE version won't provide color
			color = getDefaultBorderColor();
		}
		return color;
	}
	
	private String getFontColor(Map extraData){
		String color = (String)extraData.get("color");
		if (Strings.isEmpty(color)) {//CE version won't provide color
			color = getDefaultFontColor();
		}
		return color;
	}
	
	private String getFillColor(Map extraData){
		String color = (String)extraData.get("color");
		if (Strings.isEmpty(color)) {//CE version won't provide color
			color = getDefaultFillColor();
		}
		return color;
	}
	
	protected void doHideRow() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		range = range.getRowRange();
		CellOperationUtil.hide(range);
	}

	protected void doUnhideRow() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		range = range.getRowRange();
		CellOperationUtil.unHide(range);
	}

	protected void doUnhideColumn() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		range = range.getColumnRange();
		CellOperationUtil.hide(range);
		
	}

	protected void doHideColumn() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		range = range.getColumnRange();
		CellOperationUtil.unHide(range);
	}


// TODO, redesign this, should consider to 
//	public void toggleActionOnSheetSelected() {
//		for (Action action : toggleAction) {
//			_spreadsheet.setActionDisabled(false, action);
//		}
//		
//		//TODO: read protect information from worksheet
//		boolean protect = _spreadsheet.getSelectedXSheet().getProtect();
//		for (Action action : _defaultDisabledActionOnSheetProtected) {
//			_spreadsheet.setActionDisabled(protect, action);
//		}
//	}
//	
//	public void toggleActionOnBookClosed() {
//		for (Action a : toggleAction) {
//			_spreadsheet.setActionDisabled(true, a);
//		}
//	}
	
//	/**
//	 * Execute when user select sheet
//	 */
//	protected void doSheetSelect() {
////		syncClipboard();
////		syncAutoFilter();
//		
////		toggleActionOnSheetSelected();
//	}
	
	protected void doCloseBook() {
		Spreadsheet zss = getSpreadsheet();
		if(zss.getSrc()!=null){
			zss.setSrc(null);
		}
		if(zss.getBook()!=null){
			zss.setBook(null);
		}
		
		_clipboard = null;
//		_insertPictureSelection = null;
		
//		toggleActionOnBookClosed();
	}
	
	
	protected void doInsertPicture(){
		final Spreadsheet spreadsheet = getSpreadsheet();
		final Sheet sheet = getSheet();
		final Rect selection = getSelection();
		
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		
		askUploadFile(new EventListener<Event>() {
			public void onEvent(Event event) throws Exception {
				if(Events.ON_UPLOAD.equals(event.getName())){
					Media media = ((UploadEvent)event).getMedia();
					doInsertPicture(spreadsheet,sheet,selection,media);
				}
			}
		});
	}
	
	protected void doInsertPicture(Spreadsheet spreadsheet,Sheet sheet,Rect selection,Media media) {
		if(media==null){
			showWarnMessage("Can't get the uploaded file");
			return;
		}
		
		if(!(media instanceof AImage) || SheetOperationUtil.getPictureFormat((AImage)media)==null){
			showWarnMessage("Can't support the uploaded file");
			return;
		}
		
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		
		SheetOperationUtil.addPicture(range,(AImage)media);
		
		clearClipboard();
	}
//	protected void doInsertPicture() {
		//TODO DENNIS, re-design this.
//		if (_insertPictureSelection == null) {
//			return;
//		}
//		
//		XSheet sheet = _spreadsheet.getSelectedXSheet();
//		if (sheet != null) {
//			if (!sheet.getProtect()) {
//				final Media media = evt.getMedia();
//				if (media instanceof AImage) {
//					AImage image = (AImage)media;
//					XRanges
//					.range(_spreadsheet.getSelectedXSheet())
//					.addPicture(getClientAnchor(_insertPictureSelection.getTop(), _insertPictureSelection.getLeft(), 
//							image.getWidth(), image.getHeight()), image.getByteData(), getImageFormat(image));
//				}	
//			} else {
//				showProtectMessage();
//			}
//		}
//	}
	
//	/**
//	 * @param selection
//	 */
//	protected void doBeforeInsertPicture() {
//		_insertPictureSelection = selection;
//	}
	

	
//	protected void syncClipboard() {
//		//TODO, Dennis, shouldn't clipboard only care the range?
//		//why should we check this?
//		if (_clipboard != null) {
//			final Book srcBook = _clipboard.book;
//			if (!srcBook.equals(_spreadsheet.getBook())) {
//				_clipboard = null;
//			} else {
//				final Sheet srcSheet = _clipboard.sourceSheet;
//				boolean validSheet = srcBook.getSheetIndex(srcSheet) >= 0;
//				if (!validSheet) {
//					clearClipboard();
//				} else if (!srcSheet.equals(_spreadsheet.getSelectedXSheet())) {
//					_spreadsheet.setHighlight(null);
//				} else {
//					_spreadsheet.setHighlight(_clipboard.sourceRect);
//				}
//			}
//		}
//	}
	
//	protected void syncAutoFilter() {
//		//TODO Dennis, need to check the logic here
//		final Sheet worksheet = _spreadsheet.getTargetSheet();
//		boolean appliedFilter = false;
//		AutoFilter af = worksheet.getAutoFilter();
//		if (af != null) {
//			final CellRangeAddress afrng = af.getRangeAddress();
//			if (afrng != null) {
//				int rowIdx = afrng.getFirstRow() + 1;
//				for (int i = rowIdx; i <= afrng.getLastRow(); i++) {
//					final Row row = worksheet.getRow(i);
//					if (row != null && row.getZeroHeight()) {
//						appliedFilter = true;
//						break;
//					}
//				}	
//			}
//		}
//
//		_spreadsheet.setActionDisabled(!appliedFilter, Action.CLEAR_FILTER);
//		_spreadsheet.setActionDisabled(!appliedFilter, Action.REAPPLY_FILTER);
//		
//		if (!Objects.equals(_book, _spreadsheet.getBook())) {
//			if (_book != null) {
//				_book.unsubscribe(_bookListener);
//			}
//			_book = _spreadsheet.getXBook();
//			_book.subscribe(_bookListener);
//		}
//	}
	
//	private void init() {
//		_spreadsheet.addEventListener(Events.ON_SHEET_SELECT, _doSelectSheetListener);
//		_spreadsheet.addEventListener(Events.ON_CTRL_KEY, _doCtrlKeyListener);
//		
//		_spreadsheet.addEventListener(Events.ON_CELL_DOUBLE_CLICK, _doClearClipboard);
//		_spreadsheet.addEventListener(Events.ON_START_EDITING, _doClearClipboard);
//		
////		if (_upload == null) {
////			_upload = new Upload();
////			_upload.appendChild(_insertPicture = new Uploader());
////			_insertPicture.addEventListener(org.zkoss.zk.ui.event.Events.ON_UPLOAD, new EventListener() {
////				
////				@Override
////				public void onEvent(Event event) throws Exception {
////					doInsertPicture((UploadEvent)event);
////				}
////			});
////		}
////		_upload.setParent(_spreadsheet);
//		
//		initToggleAction();
//	}
	
	/**
	 * Execute when user press key
	 * @param event
	 */
	protected void doCtrlKey(KeyEvent event) {
		Rect selection = getSelection();
		if (46 == event.getKeyCode()) {
			if (event.isCtrlKey())
				doClearStyle();
			else
				doClearContent();
			return;
		}
		if (false == event.isCtrlKey())
			return;
		
		char keyCode = (char) event.getKeyCode();
		switch (keyCode) {
		case 'X':
			doCut();
			break;
		case 'C':
			doCopy();
			break;
		case 'V':
			if (_clipboard != null){
				//what the god-dam implementation in user customizable side here
				//TODO
//				_spreadsheet.smartUpdate("doPasteFromServer", true);
			}
			doPaste();
			break;
		case 'D':
			doClearContent();
			break;
		case 'B':
			doFontBold();
			break;
		case 'I':
			doFontItalic();
			break;
		case 'U':
			doFontUnderline();
			break;
		}
	}
	
//	/**
//	 * Bind the handler's target
//	 * 
//	 * @param spreadsheet
//	 */
//	public void bind(Spreadsheet spreadsheet) {
//		if (_spreadsheet != spreadsheet) {
//			_spreadsheet = spreadsheet;
//			init();	
//		}
//	}
	
//	/**
//	 * Unbind the handler's target
//	 */
//	public void unbind() {
//		if (_spreadsheet != null) {
////			if (_upload.getParent() == _spreadsheet) {
////				_spreadsheet.removeChild(_upload);
////			}
//			
//			_spreadsheet.removeEventListener(Events.ON_SHEET_SELECT, _doSelectSheetListener);
//			_spreadsheet.removeEventListener(Events.ON_CTRL_KEY, _doCtrlKeyListener);
//		}
//	}
	
	
	protected Clipboard getClipboard() {
		return _clipboard;
	}
	
	protected void setClipboard(Clipboard clipboard){
		_clipboard = clipboard;
	}
	
	protected void clearClipboard() {
		_clipboard = null;
		getSpreadsheet().setHighlight(null);
		//TODO: shall also clear client side clipboard if possible
	}
	
	/**
	 * Execute when user click copy
	 */
	protected void doCopy() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		setClipboard(new Clipboard(Clipboard.Type.COPY, sheet.getSheetName(), selection));
		getSpreadsheet().setHighlight(getSelection());
	}
	
	protected void doPaste(PasteType pasteType, PasteOperation pasteOperation, boolean skipBlank, boolean transpose) {
		Clipboard cb = getClipboard();
		if(cb==null)
			return;
		
		Book book = getBook();
		Sheet destSheet = getSheet();
		Sheet srcSheet = book.getSheet(cb.sourceSheetName);
		if(srcSheet==null){
			//TODO message;
			clearClipboard();
			return;
		}
		Rect src = cb.sourceRect;
		
		Rect selection = getSelection();
		
		Range srcRange = Ranges.range(srcSheet, src.getTop(),
				src.getLeft(), src.getBottom(),src.getRight());

		Range destRange = Ranges.range(destSheet, selection.getTop(),
				selection.getLeft(), selection.getBottom(), selection.getRight());
		
		if (destRange.isProtected()) {
			showProtectMessage();
			return;
		} else if (cb.type == Type.CUT && srcRange.isProtected()) {
			showProtectMessage();
			return;
		}
		
		if(cb.type==Type.CUT){
			CellOperationUtil.cut(srcRange,destRange);
			clearClipboard();
		}else{
			CellOperationUtil.pasteSpecial(srcRange, destRange, pasteType, pasteOperation, skipBlank, transpose);
		}
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
	
	/**
	 * Execute when user click paste 
	 */
	protected void doPaste() {
		doPaste(PasteType.PASTE_ALL,PasteOperation.PASTEOP_NONE,false,false);
	}
	
	
	/**
	 * Execute when user click paste formula
	 */
	protected void doPasteFormula() {
		doPaste(PasteType.PASTE_FORMULAS,PasteOperation.PASTEOP_NONE,false,false);
	}
	
	/**
	 *  Execute when user click paste value
	 */
	protected void doPasteValue() {
		doPaste(PasteType.PASTE_VALUES,PasteOperation.PASTEOP_NONE,false,false);
	}
	
	/**
	 * Execute when user click paste all except border
	 */
	protected void doPasteAllExceptBorder() {
		doPaste(PasteType.PASTE_ALL_EXCEPT_BORDERS,PasteOperation.PASTEOP_NONE,false,false);
	}
	
	/**
	 * Execute when user click paste transpose
	 */
	protected void doPasteTranspose() {
		doPaste(PasteType.PASTE_ALL, PasteOperation.PASTEOP_NONE, false, true);
	}
	
	/**
	 * Execute when user click cut
	 */
	protected void doCut() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		
		setClipboard(new Clipboard(Clipboard.Type.CUT, sheet.getSheetName(), selection));
		getSpreadsheet().setHighlight(getSelection());
	}
	
	/**
	 * Execute when user select font family
	 * 
	 * @param selection
	 */
	protected void doFontFamily(String fontFamily) {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		CellOperationUtil.applyFontName(range, fontFamily);
	}
	
	/**
	 * Execute when user select font size
	 *  
	 * @param selection
	 */
	protected void doFontSize(int fontSize) {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		CellOperationUtil.applyFontSize(range, (short)fontSize);
	}
	
	/**
	 * Execute when user click font bold
	 * 
	 * @param selection
	 */
	protected void doFontBold() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}

		//toggle and apply bold of first cell to dest
		Boldweight bw = range.getCellStyle().getFont().getBoldweight();
		if(Boldweight.BOLD.equals(bw)){
			bw = Boldweight.NORMAL;
		}else{
			bw = Boldweight.BOLD;
		}
		
		CellOperationUtil.applyFontBoldweight(range, bw);
	}
	
	/**
	 * Execute when user click font italic
	 * 
	 * @param selection
	 */
	protected void doFontItalic() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}

		//toggle and apply bold of first cell to dest
		boolean italic = !range.getCellStyle().getFont().isItalic();
		CellOperationUtil.applyFontItalic(range, italic);	
	}
	
	/**
	 * Execute when user click font strikethrough
	 * 
	 * @param selection
	 */
	protected void doFontStrikeout() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}

		//toggle and apply bold of first cell to dest
		boolean strikeout = !range.getCellStyle().getFont().isStrikeout();
		CellOperationUtil.applyFontStrikeout(range, strikeout);
	}
	
	/**
	 * Execute when user click font underline
	 * 
	 * @param selection
	 */
	protected void doFontUnderline() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}

		//toggle and apply bold of first cell to dest
		Underline underline = range.getCellStyle().getFont().getUnderline();
		if(Underline.NONE.equals(underline)){
			underline = Underline.SINGLE;
		}else{
			underline = Underline.NONE;
		}
		
		CellOperationUtil.applyFontUnderline(range, underline);	
	}
	
	
	protected void doBorder(ApplyBorderType type,BorderType borderTYpe, String color){
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		CellOperationUtil.applyBorder(range,type, borderTYpe, color);
	}
	
	/**
	 * Execute when user click border
	 * 
	 * @param selection
	 */
	protected void doBorder(String color) {
		doBorder(ApplyBorderType.EDGE_BOTTOM,BorderType.MEDIUM,color);
	}
	
	/**
	 * Execute when user click bottom border
	 * 
	 * @param selection
	 */
	protected void doBorderBottom(String color) {
		doBorder(ApplyBorderType.EDGE_BOTTOM,BorderType.MEDIUM,color);
	}
	
	/**
	 * Execute when user click top border
	 * 
	 * @param selection
	 */
	protected void doBoderTop(String color) {
		doBorder(ApplyBorderType.EDGE_TOP,BorderType.MEDIUM,color);
	}
	
	/**
	 * Execute when user click left border
	 * 
	 * @param selection
	 */
	protected void doBorderLeft(String color) {
		doBorder(ApplyBorderType.EDGE_LEFT,BorderType.MEDIUM,color);
	}
	
	/**
	 * Execute when user click right border
	 * 
	 * @param selection
	 */
	protected void doBorderRight(String color) {
		doBorder(ApplyBorderType.EDGE_RIGHT,BorderType.MEDIUM,color);
	}
	
	/**
	 * Execute when user click no border
	 * 
	 * @param selection
	 */
	protected void doBorderNo(String color) {
		doBorder(ApplyBorderType.FULL,BorderType.NONE,color);
	}
	
	/**
	 * Execute when user click all border
	 * 
	 * @param selection
	 */
	protected void doBorderAll(String color) {
		doBorder(ApplyBorderType.FULL,BorderType.MEDIUM,color);
	}
	
	/**
	 * Execute when user click outside border
	 * 
	 * @param selection
	 */
	protected void doBorderOutside(String color) {
		doBorder(ApplyBorderType.OUTLINE,BorderType.MEDIUM,color);
	}
	
	/**
	 * Execute when user click inside border
	 * 
	 * @param selection
	 */
	protected void doBorderInside(String color) {
		doBorder(ApplyBorderType.INSIDE,BorderType.MEDIUM,color);
	}
	
	/**
	 * Execute when user click inside horizontal border
	 * 
	 * @param selection
	 */
	protected void doBorderInsideHorizontal(String color) {
		doBorder(ApplyBorderType.INSIDE_HORIZONTAL,BorderType.MEDIUM,color);
	}
	
	/**
	 * Execute when user click inside vertical border
	 * 
	 * @param selection
	 */
	protected void doBorderInsideVertical(String color) {
		doBorder(ApplyBorderType.INSIDE_VERTICAL,BorderType.MEDIUM,color);
	}
	
	/**
	 * Execute when user click font color 
	 * 
	 * @param color
	 * @param selection
	 */
	protected void doFontColor(String color) {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		CellOperationUtil.applyFontColor(range, color);	
	}
	
	/**
	 * Execute when user click fill color
	 * 
	 * @param color
	 * @param selection
	 */
	protected void doFillColor(String color) {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		CellOperationUtil.applyCellColor(range,color);
	}
	
	/**
	 * Execute when user click vertical align top
	 * 
	 * @param selection
	 */
	protected void doVerticalAlignTop() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		CellOperationUtil.applyCellVerticalAlignment(range, VerticalAlignment.TOP);
	}
	
	/**
	 * Execute when user click vertical align middle
	 * 
	 * @param selection
	 */
	protected void doVerticalAlignMiddle() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		CellOperationUtil.applyCellVerticalAlignment(range, VerticalAlignment.CENTER);
	}

	/**
	 * Execute when user click vertical align bottom
	 * 
	 * @param selection
	 */
	protected void doVerticalAlignBottom() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		CellOperationUtil.applyCellVerticalAlignment(range, VerticalAlignment.BOTTOM);
	}
	
	/**
	 * @param selection
	 */
	protected void doHorizontalAlignLeft() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		CellOperationUtil.applyCellAlignment(range, Alignment.LEFT);
	}
	
	/**
	 * @param selection
	 */
	protected void doHorizontalAlignCenter() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		CellOperationUtil.applyCellAlignment(range, Alignment.CENTER);
	}
	
	/**
	 * @param selection
	 */
	protected void doHorizontalAlignRight() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		CellOperationUtil.applyCellAlignment(range, Alignment.RIGHT);
	}
	
	/**
	 * @param selection
	 */
	protected void doWrapText() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		boolean wrapped = !range.getCellStyle().isWrapText();
		CellOperationUtil.applyCellWrapText(range, wrapped);
	}
	
	/**
	 * @param selection
	 */
	protected void doMergeAndCenter() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		CellOperationUtil.toggleMergeCenter(range);
		clearClipboard();
	}
	
	/**
	 * @param selection
	 */
	protected void doMergeAcross() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		CellOperationUtil.merge(range, true);
		clearClipboard();
	}
	
	/**
	 * @param selection
	 */
	protected void doMergeCell() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		CellOperationUtil.merge(range, false);
		clearClipboard();
	}
	
	/**
	 * @param selection
	 */
	protected void doUnmergeCell() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		CellOperationUtil.unMerge(range);
		clearClipboard();
	}
	
	/**
	 * @param selection
	 */
	protected void doShiftCellRight() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		CellOperationUtil.insert(range,InsertShift.RIGHT, InsertCopyOrigin.RIGHT_BELOW);
		clearClipboard();
	}
	
	/**
	 * @param selection
	 */
	protected void doShiftCellDown() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		CellOperationUtil.insert(range,InsertShift.DOWN, InsertCopyOrigin.LEFT_ABOVE);
		clearClipboard();
	}
	
	/**
	 * @param selection
	 */
	protected void doInsertSheetRow() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		
		if(range.isWholeColumn()){
			showWarnMessage("don't allow to inser row when select whole column");
			return;
		}
		
		range = range.getRowRange();
		CellOperationUtil.insert(range,InsertShift.DOWN, InsertCopyOrigin.LEFT_ABOVE);
		clearClipboard();
		
	}
	
	/**
	 * @param selection
	 */
	protected void doInsertSheetColumn() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		
		if(range.isWholeRow()){
			showWarnMessage("don't allow to inser column when select whole row");
			return;
		}
		
		range = range.getColumnRange();
		CellOperationUtil.insert(range,InsertShift.RIGHT, InsertCopyOrigin.RIGHT_BELOW);
		clearClipboard();
	}
	
	/**
	 * @param selection
	 */
	protected void doShiftCellLeft() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		
		CellOperationUtil.delete(range,DeleteShift.LEFT);
		clearClipboard();
	}
	
	/**
	 * @param selection
	 */
	protected void doShiftCellUp() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		
		CellOperationUtil.delete(range,DeleteShift.UP);
		clearClipboard();
	}
	
	/**
	 * @param selection
	 */
	protected void doDeleteSheetRow() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		
		if(range.isWholeColumn()){
			showWarnMessage("don't allow to delete all rows");
			return;
		}
		
		range = range.getRowRange();
		CellOperationUtil.delete(range, DeleteShift.UP);
		clearClipboard();
	}
	
	/**
	 * @param selection
	 */
	protected void doDeleteSheetColumn() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		
		if(range.isWholeRow()){
			showWarnMessage("don't allow to delete all column");
			return;
		}
		
		range = range.getColumnRange();
		CellOperationUtil.delete(range, DeleteShift.LEFT);
		clearClipboard();
	}

	/**
	 * Execute when user click clear style
	 */
	protected void doClearStyle() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		CellOperationUtil.clearStyles(range);
	}
	
	/**
	 * Execute when user click clear content
	 */
	protected void doClearContent() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		CellOperationUtil.clearContents(range);
	}
	
	/**
	 * Execute when user click clear all
	 */
	protected void doClearAll() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		CellOperationUtil.clearAll(range);
	}
	
	/**
	 * @param selection
	 */
	protected void doSortAscending() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		CellOperationUtil.sort(range,false);
		clearClipboard();
	}
	
	/**
	 * @param selection
	 */
	protected void doSortDescending() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		CellOperationUtil.sort(range,true);
		clearClipboard();
	}
	
	/**
	 * @param selection
	 */
	protected void doFilter() {
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		SheetOperationUtil.toggleAutoFilter(range);
		clearClipboard();
	}
	
	/**
	 * @param selection
	 */
	protected void doClearFilter() {
		Sheet sheet = getSheet();
		
		Range range = Ranges.range(sheet);
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		SheetOperationUtil.resetAutoFilter(range);
		clearClipboard();
	}
	
	/**
	 * @param selection
	 */
	protected void doReapplyFilter() {
		Sheet sheet = getSheet();
		
		Range range = Ranges.range(sheet);
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		SheetOperationUtil.applyAutoFilter(range);
		clearClipboard();
	}
	
	/**
	 * 
	 */
	protected void doProtectSheet() {
		
		Sheet sheet = getSheet();
		
		Range range = Ranges.range(sheet);
		
		String newpassword = "1234";//TODO, make it meaningful
		if(range.isProtected()){
			SheetOperationUtil.protectSheet(range,null,null);
		}else{
			SheetOperationUtil.protectSheet(range,null,newpassword);
		}
		
		boolean p = range.isProtected();
		
		//TODO re-factor action bar
//		for (Action action : _defaultDisabledActionOnSheetProtected) {
//			getSpreadsheet().setActionDisabled(p, action);
//		}
	}
	
	/**
	 * 
	 */
	protected void doGridlines() {
		Sheet sheet = getSheet();
		
		Range range = Ranges.range(sheet);
		
		SheetOperationUtil.displaySheetGridlines(range,!range.isDisplaySheetGridlines());
	}
	
	
	protected void doChart(Chart.Type type, Chart.Grouping grouping, Chart.LegendPosition pos){
		Sheet sheet = getSheet();
		Rect selection = getSelection();
		Range range = Ranges.range(sheet, selection.getTop(), selection.getLeft(), selection.getBottom(), selection.getRight());
		if(range.isProtected()){
			showProtectMessage();
			return;
		}
		
		SheetAnchor anchor = SheetOperationUtil.toChartAnchor(range);
		
		ChartData data = org.zkoss.zss.api.ChartDataUtil.getChartData(sheet,selection, type);
		SheetOperationUtil.addChart(range,anchor,data,type,grouping,pos);
		clearClipboard();
		
	}
	
	/**
	 * @param selection
	 */
	protected void doColumnChart() {
		doChart(Chart.Type.Column,Chart.Grouping.STANDARD, Chart.LegendPosition.RIGHT);
	}
	
	/**
	 * @param selection
	 */
	protected void doColumnChart3D() {
		doChart(Chart.Type.Column3D,Chart.Grouping.STANDARD, Chart.LegendPosition.RIGHT);
	}
	/**
	 * @param selection
	 */
	protected void doLineChart() {
		doChart(Chart.Type.Line,Chart.Grouping.STANDARD, Chart.LegendPosition.RIGHT);
	}
	/**
	 * @param selection
	 */
	protected void doLineChart3D() {
		doChart(Chart.Type.Line3D,Chart.Grouping.STANDARD, Chart.LegendPosition.RIGHT);
	}
	/**
	 * @param selection
	 */
	protected void doPieChart() {
		doChart(Chart.Type.Pie,Chart.Grouping.STANDARD, Chart.LegendPosition.RIGHT);
	}
	/**
	 * @param selection
	 */
	protected void doPieChart3D() {
		doChart(Chart.Type.Pie3D,Chart.Grouping.STANDARD, Chart.LegendPosition.RIGHT);
	}
	/**
	 * @param selection
	 */
	protected void doBarChart() {
		doChart(Chart.Type.Bar,Chart.Grouping.STANDARD, Chart.LegendPosition.RIGHT);
	}
	/**
	 * @param selection
	 */
	protected void doBarChart3D() {
		doChart(Chart.Type.Bar3D,Chart.Grouping.STANDARD, Chart.LegendPosition.RIGHT);
	}
	/**
	 * @param selection
	 */
	protected void doAreaChart() {
		doChart(Chart.Type.Area,Chart.Grouping.STANDARD, Chart.LegendPosition.RIGHT);
	}
	/**
	 * @param selection
	 */
	protected void doScatterChart() {
		doChart(Chart.Type.Scatter,Chart.Grouping.STANDARD, Chart.LegendPosition.RIGHT);
	}
	/**
	 * @param selection
	 */
	protected void doDoughnutChart() {
		doChart(Chart.Type.Doughnut,Chart.Grouping.STANDARD, Chart.LegendPosition.RIGHT);
	}
	
	private static <T> T checkNotNull(String message, T t) {
		if (t == null) {
			throw new NullPointerException(message);
		}
		return t;
	}

	private static class UserActionContext {

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
	 * Used for copy & paste function
	 * 
	 * @author sam
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
	
	protected void doRowHeight() {
		showNotImplement(Action.ROW_HEIGHT.toString());
	}

	protected  void doColumnWidth() {
		showNotImplement(Action.COLUMN_WIDTH.toString());
	}

	protected  void doFormatCell() {
		showNotImplement(Action.FORMAT_CELL.toString());
	}

	protected  void doHyperlink() {
		showNotImplement(Action.HYPERLINK.toString());
	}

	protected  void doCustomSort() {
		showNotImplement(Action.CUSTOM_SORT.toString());
	}

	protected  void doPasteSpecial() {
		showNotImplement(Action.PASTE_SPECIAL.toString());
	}

	protected  void doExportPDF() {
		showNotImplement(Action.EXPORT_PDF.toString());
	}

	protected  void doSaveBook() {
		showNotImplement(Action.SAVE_BOOK.toString());
	}

	protected void doNewBook() {
		showNotImplement(Action.NEW_BOOK.toString());
	}

	protected void doShowFormulaPanel() {
		showNotImplement(Action.FORMULA_PANEL.toString());
	}

	protected void doShowInsertPanel() {
		showNotImplement(Action.INSERT_PANEL.toString());
	}

	protected  void doShowHomePanel() {
		showNotImplement(Action.HOME_PANEL.toString());
	}

	protected void doInsertFunction() {
		showNotImplement(Action.INSERT_FUNCTION.toString());
	}
	
	protected void askUploadFile(EventListener l){
		//TODO Need ZK's new feature support
		
	}
}
