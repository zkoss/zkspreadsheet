/* MainWindow.java

{{IS_NOTE
	Purpose:

	Description:

	History:
		Dec 20, 2007 12:40:46 PM     2007, Created by Dennis.Chenf
		Jan 1, 2009 modified by kinda lu
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

 */
package org.zkoss.zss.app;

import java.util.Iterator;

import org.zkoss.lang.Library;
//import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.Path;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zss.app.ctrl.HeaderSizeCtrl;
import org.zkoss.zss.app.ctrl.RenameSheetCtrl;
import org.zkoss.zss.app.file.FileHelper;
import org.zkoss.zss.app.formula.FormulaMetaInfo;
import org.zkoss.zss.app.zul.Dialog;
import org.zkoss.zss.app.zul.FileMenu;
import org.zkoss.zss.app.zul.ViewMenu;
import org.zkoss.zss.app.zul.Zssapp;
import org.zkoss.zss.app.zul.Zssapps;
import org.zkoss.zss.app.zul.ctrl.DesktopCellStyleContext;
import org.zkoss.zss.app.zul.ctrl.DesktopWorkbenchContext;
import org.zkoss.zss.app.zul.ctrl.SSRectCellStyle;
import org.zkoss.zss.app.zul.ctrl.SSWorkbookCtrl;
import org.zkoss.zss.app.zul.ctrl.WorkbenchCtrl;
import org.zkoss.zss.app.zul.ctrl.WorkbookCtrl;
import org.zkoss.zss.engine.event.SSDataEvent;
//import org.zkoss.zss.model.sys.XRange;
//import org.zkoss.zss.model.sys.XRanges;
//import org.zkoss.zss.model.sys.XSheet;
//import org.zkoss.zss.model.sys.impl.BookHelper;
import org.zkoss.zss.ui.UserAction;
import org.zkoss.zss.ui.Position;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.KeyEvent;
//import org.zkoss.zss.ui.impl.MergeMatrixHelper;
//import org.zkoss.zss.ui.impl.MergedRect;
//import org.zkoss.zss.ui.impl.Utils;
//import org.zkoss.zss.ui.sys.XActionHandler;
import org.zkoss.zul.Combobox;
import org.zkoss.zul.Div;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.Window;

/**
 * @author Peter Kuo
 * @modify kinda lu
 */
public class MainWindowCtrl extends GenericForwardComposer implements WorkbenchCtrl {

	private static final long serialVersionUID = 1;
	
	private final static String KEY_PDF = "org.zkoss.zss.app.exportToPdf";
	private final static String KEY_HTML = "org.zkoss.zss.app.exportToHtml";
	
	Spreadsheet spreadsheet;

	RangeHelper rangeh;

	//Window mainWin;
	Div mainWin;
	/*Menus*/
	FileMenu fileMenu;
	ViewMenu viewMenu;
	
	/*Dialog*/
	Dialog _insertFormulaDialog;
	Dialog _insertHyperlinkDialog;
	Dialog _pasteSpecialDialog;
	Dialog _composeFormulaDialog;
	Dialog _formatNumberDialog;
	Dialog _saveFileDialog;
	Dialog _headerSizeDialog;
	Dialog _renameSheetDialog;
	Dialog _openFileDialog;
	Dialog _importFileDialog;
	Dialog _customSortDialog;
	Dialog _exportToPdfDialog;
	Dialog _exportToHtmlDialog;
	
	MainActionHandler actionHandler;

	public Spreadsheet getSpreadsheet() {
		return spreadsheet;
	}

	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		spreadsheet.setUserActionHandler(actionHandler = new MainActionHandler());
		
		//TODO: do it after "afterCompose"
		FileHelper.openNewSpreadsheet(spreadsheet);
		
		//TODO: replace this mechanism
		initZssappComponents();
		init();
		rangeh = new RangeHelper(spreadsheet);
	}
	
	public void onCreate() {
		getDesktopWorkbenchContext().fireWorkbookChanged();
	}
	
	//TODO: remove this mechanism
	private void initZssappComponents() {
		Zssapps.bindSpreadsheet(spreadsheet, this);
	}
	
	public void init() {
		boolean isPE = WebApps.getFeature("pe");
		
		//Note. setSrcName will set spreadsheet's src name, but not the book name
		// if setSrc will init a book, then setSrcName only change the src name, 
		// if setSrc again with the first same book, the book will register two listener 
		//spreadsheet.setSrcName("Untitled");
		final DesktopWorkbenchContext workbenchContext = getDesktopWorkbenchContext();
		workbenchContext.doTargetChange(new SSWorkbookCtrl(spreadsheet));
		workbenchContext.setWorkbenchCtrl(this);
		
//		if (!FileHelper.hasSavePermission())
//			saveBtn.setVisible(false);
		workbenchContext.addEventListener(Consts.ON_WORKBOOK_SAVED,	new EventListener() {
			public void onEvent(Event event) throws Exception {
				if (!FileHelper.hasSavePermission())
					return;
				
				spreadsheet.setActionDisabled(true, UserAction.SAVE_BOOK);
			}
		});

		workbenchContext.addEventListener(Consts.ON_WORKBOOK_CHANGED, new EventListener() {
			public void onEvent(Event event) throws Exception {
				boolean isOpen = spreadsheet.getXBook() != null;
//				toolbarMask.setVisible(!isOpen);
//				closeBtn.setVisible(isOpen);
				
				spreadsheet.setActionDisabled(true, UserAction.SAVE_BOOK);

//				gridlinesCheckbox.setChecked(isOpen && spreadsheet.getSelectedSheet().isDisplayGridlines());
//				protectSheet.setChecked(isOpen && spreadsheet.getSelectedSheet().getProtect());
//				protectSheet.setDisabled(!isOpen);
				
				//TODO: provide clip board interface, to allow save cut, copy, high light info
				//use set setHighlight null can cancel selection, but need to re-store selection when select same sheet again
				spreadsheet.setHighlight(null);
				
				if (isOpen) {
					getCellStyleContext().doTargetChange(
							new SSRectCellStyle(Utils.getOrCreateCell(spreadsheet.getSelectedXSheet(), 0, 0), spreadsheet));
//					syncAutoFilterStatus();
				}
			}
		});

		workbenchContext.addEventListener(Consts.ON_SHEET_CONTENTS_CHANGED,  new EventListener(){
			public void onEvent(Event event) throws Exception {
				doContentChanged();
			}}
		);
		
		workbenchContext.addEventListener(Consts.ON_SHEET_INSERT_FORMULA, new EventListener() {
			public void onEvent(Event event) throws Exception {
				String formula = (String)event.getData();
				Rect rect = spreadsheet.getSelection();
				XRange rng = XRanges.range(spreadsheet.getSelectedXSheet(), rect.getTop(), rect.getLeft());
				rng.setEditText(formula);
			}
		});
		workbenchContext.getWorkbookCtrl().addBookEventListener(new EventListener() {
			public void onEvent(Event event) throws Exception {
				String evtName = event.getName();
				if (evtName == SSDataEvent.ON_CONTENTS_CHANGE) {
					doContentChanged();
				}
			}
		});
	}

	public void doContentChanged() {
		//enable SAVE_BOOK button
		 boolean savePermission = FileHelper.hasSavePermission();
		 if (savePermission) {
			 spreadsheet.setActionDisabled(false, UserAction.SAVE_BOOK);
		 }
		
		XSheet seldSheet = spreadsheet.getSelectedXSheet();
		Rect seld =  spreadsheet.getSelection();
		int row = seld.getTop();
		int col = seld.getLeft();
		Cell cell = Utils.getCell(seldSheet, row, col);
		if (cell != null) {
			getCellStyleContext().doTargetChange(new SSRectCellStyle(cell, spreadsheet));
		}
	}
	
	public void saveBook() {
		DesktopWorkbenchContext workbench = getDesktopWorkbenchContext();
		if (workbench.getWorkbookCtrl().hasFileExtentionName()) {
			workbench.getWorkbookCtrl().save();
			workbench.fireWorkbookSaved();
		} else
			workbench.getWorkbenchCtrl().openSaveFileDialog();
	}

//	public void syncAutoFilterStatus() {
//		final Worksheet worksheet = spreadsheet.getSelectedSheet();
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
////		clearFilter.setDisabled(!appliedFilter);
////		reapplyFilter.setDisabled(!appliedFilter);
//	}
	
	private WorkbookCtrl getWorkbookCtrl() {
		return  getDesktopWorkbenchContext().getWorkbookCtrl();
	}
	
	private DesktopWorkbenchContext getDesktopWorkbenchContext() {
		return Zssapp.getDesktopWorkbenchContext(self);
	}
	
//	private boolean isMergedCell(int tRow, int lCol, int bRow, int rCol) {
//		MergeMatrixHelper mmhelper = spreadsheet.getMergeMatrixHelper(spreadsheet.getSelectedXSheet());
//		for (final Iterator iter = mmhelper.getRanges().iterator(); iter
//				.hasNext();) {
//			MergedRect block = (MergedRect) iter.next();
//			int bt = block.getTop();
//			int bl = block.getLeft();
//			int bb = block.getBottom();
//			int br = block.getRight();
//			if (lCol <= bl && tRow <= bt && rCol >= br && bRow >= bb) {
//				return true;
//			}
//		}
//		return false;
//	}

	/**
	 * @return
	 */
	private DesktopCellStyleContext getCellStyleContext() {
		return Zssapp.getDesktopCellStyleContext(self);
	}

	public void onRevision(ForwardEvent event) {
		throw new UiException("revision not implement yet");
//		reloadRevisionMenu();
//		Window win = (Window) mainWin.getFellow("revisionWin");
//		win.setPosition("parent");
//		win.setLeft("250px");
//		win.setTop("250px");
//		win.doPopup();
	}

	// SECTION EDIT MENU
	public void onEditUndo(ForwardEvent event) {
		throw new UiException("Undo not implmented yet");
		// spreadsheet.undo();
	}

	public void onEditRedo(ForwardEvent event) {
		throw new UiException("Redo not implemented yet");
		// spreadsheet.redo();
	}

	// SECTION FORMAT MENU
//	public void onFormatNumber(ForwardEvent event) {
//		Window win = (Window) mainWin.getFellow("formatNumberWin");
//		win.setPosition("parent");
//		win.setLeft("170px");
//		win.setTop("24px");
//		win.doPopup();// Modal();
//	}

	// SECTION HELP MENU
	public void onHelpCheatsheet(ForwardEvent event) {
		Window win = (Window) mainWin.getFellow("cheatsheet");
		win.setPosition("parent");
		win.setLeft("327px");
		win.setTop("124px");
		win.doPopup();
	}

	public void onRange(ForwardEvent event) {
		rangeh.dispatcher((String) event.getData());
	}

	public void reloadRevisionMenu() {
		throw new UiException("reloadRevision not implement yet");
//		try {
//			// read the metafile
//			BufferedReader bReader;
//			bReader = new BufferedReader(new FileReader(fileh.xlsDir
//					+ "metaFile"));
//			String line = null;
//			Stack stack = new Stack();
//
//			String timeStr, filename, hashFilename;
//			while (true) {
//				timeStr = bReader.readLine();
//				if (timeStr == null) {
//					break;
//				}
//				filename = bReader.readLine();
//				if (filename == null) {
//					System.out.println("Warning: filename cannot read from metaFile");
//					break;
//				}
//				hashFilename = bReader.readLine();
//				if (hashFilename == null) {
//					System.out.println("Warning: hashFilename cannot read from metaFile");
//					break;
//				}
//				if (spreadsheet.getSrc().equals(filename)) {
//					stack.add(timeStr);
//					stack.add(filename);
//					stack.add(hashFilename);
//				}
//			}
//			bReader.close();
//
//			// put read data to Menu
//			org.zkoss.zul.Row newRow;
//			Rows revisionRows = (Rows) mainWin.getFellow("revisionWin")
//					.getFellow("revisionRows");
//			List rowsChildList = revisionRows.getChildren();
//			while (!rowsChildList.isEmpty())
//				revisionRows.removeChild((Component) rowsChildList.get(0));
//			while (!stack.isEmpty()) {
//
//				hashFilename = (String) stack.pop();
//				filename = (String) stack.pop();
//				timeStr = (String) stack.pop();
//				String dateStr = new Date(Long.parseLong(timeStr)).toString();
//
//				newRow = new org.zkoss.zul.Row();
//				Radio radio = new Radio();
//				radio.setAttribute("value", hashFilename);
//
//				if (hashFilename.equals(this.hashFilename)) {
//					newRow.appendChild(new Label("current"));
//					newRow.setStyle("background:rgb(250,230,180) none");
//				} else {
//					newRow.appendChild(radio);
//					// p("set hashFilename: "+hashFilename);
//				}
//				newRow.appendChild(new Label(dateStr));
//				newRow.appendChild(new Label("test name"));
//				newRow.appendChild(new Label("no comment"));
//
//				revisionRows.appendChild(newRow);
//				// p(""+radio.getAttribute("value"));
//			}
//			// close the jdbc connection
//
//		} catch (Exception ex) {
//			throw new RuntimeException(ex);
//		}
	}

//	public void onRevisionOK(ForwardEvent event) {
//		String filename = null;
//
//		Rows revisionRows = (Rows) mainWin.getFellow("revisionWin").getFellow(
//				"revisionRows");
//		List rowList = revisionRows.getChildren();
//		for (int i = 0; i < rowList.size(); i++) {
//			Component tmpComponent = ((Component) rowList.get(i))
//					.getFirstChild();
//			if (tmpComponent instanceof Radio
//					&& ((Radio) tmpComponent).isChecked()) {
//				filename = (String) ((Radio) tmpComponent)
//						.getAttribute("value");
//				// p("onRevision: "+filename);
//			}
//		}
//
//		openFileInFS(filename);
//
//		Window revisionWin = (Window) Path
//				.getComponent("//p1/mainWin/revisionWin");
//		revisionWin.setVisible(false);
//		spreadsheet.invalidate();
//		// spreadsheet.notifyRevision();
//	}

//	private Connection getDBConnection() {
//		try {
//			return DriverManager.getConnection(
//					"jdbc:mysql://localhost:3306/zss", "root", "rootzk");
//		} catch (SQLException e) {
//			e.printStackTrace();
//		}
//		return null;
//	}

	public void openCustomSortDialog(Rect selection) {
		if (_customSortDialog == null || _customSortDialog.isInvalidated())
			_customSortDialog = (Dialog) Executions.createComponents(Consts._CustomSortDialog_zul, mainWin, CustomSortWindowCtrl.newArg(spreadsheet));
		_customSortDialog.fireOnOpen(selection);
	}
	
	public void openHyperlinkDialog(Rect selection) {
		if (_insertHyperlinkDialog == null || _insertHyperlinkDialog.isInvalidated())
			_insertHyperlinkDialog = (Dialog)Executions.createComponents(Consts._InsertHyperlinkDialog_zul, mainWin, Zssapps.newSpreadsheetArg(spreadsheet));
		_insertHyperlinkDialog.fireOnOpen(selection);
	}

	public void openInsertFormulaDialog(Rect selection) {
		if (_insertFormulaDialog == null || _insertFormulaDialog.isInvalidated())
			_insertFormulaDialog = (Dialog)Executions.createComponents(Consts._InsertFormulaDialog2_zul, mainWin, null);
		_insertFormulaDialog.fireOnOpen(selection);
	}

	public void openPasteSpecialDialog() {
		if (_pasteSpecialDialog == null || _pasteSpecialDialog.isInvalidated())
			_pasteSpecialDialog = (Dialog)Executions.createComponents(Consts._PasteSpecialDialog_zul, mainWin, Zssapps.newSpreadsheetArg(spreadsheet));
		_pasteSpecialDialog.fireOnOpen(null);
	}

	public boolean toggleFormulaBar() {
		spreadsheet.setShowFormulabar(!spreadsheet.isShowFormulabar());
		return spreadsheet.isShowFormulabar();
	}

	public void openComposeFormulaDialog(FormulaMetaInfo metainfo) {
		if (_composeFormulaDialog == null || _composeFormulaDialog.isInvalidated())
			_composeFormulaDialog = (Dialog) Executions.createComponents(Consts._ComposeFormulaDialog_zul, mainWin, arg);
		_composeFormulaDialog.fireOnOpen(metainfo);
	}

	public void openFormatNumberDialog(Rect selection) {
		if (_formatNumberDialog == null || _formatNumberDialog.isInvalidated())
			_formatNumberDialog = (Dialog) Executions.createComponents(Consts._FormatNumberDialog_zul, mainWin, Zssapps.newSpreadsheetArg(spreadsheet));
		_formatNumberDialog.fireOnOpen(selection);	
	}

	public void openSaveFileDialog() {
		if (_saveFileDialog == null || _saveFileDialog.isInvalidated())
			_saveFileDialog = (Dialog) Executions.createComponents(Consts._SaveFile_zul, mainWin, null);
		_saveFileDialog.fireOnOpen(null);
	}

	public void openModifyHeaderSizeDialog(int headerType, Rect selection) {
		int prev = -1;
		boolean sameVal = true;
		if (headerType == WorkbookCtrl.HEADER_TYPE_ROW) {
			for (int i = selection.getTop(); i <= selection.getBottom(); i++) {
				if (prev < 0)
					prev = Utils.getRowHeightInPx(spreadsheet.getSelectedXSheet(), i);
				else if (prev != Utils.getRowHeightInPx(spreadsheet.getSelectedXSheet(), i)) {
					sameVal = false;
					break;
				}
			}
		} else {
			for (int i = selection.getLeft(); i <= selection.getRight(); i++) {
				if (prev < 0)
					prev = Utils.getColumnWidthInPx(spreadsheet.getSelectedXSheet(), i);
				else if (prev != Utils.getColumnWidthInPx(spreadsheet.getSelectedXSheet(), i)) {
					sameVal = false;
					break;
				}
			}
		}

		if (_headerSizeDialog == null || _headerSizeDialog.isInvalidated())
			_headerSizeDialog = (Dialog) Executions.createComponents(Consts._HeaderSize_zul, mainWin, null);
		_headerSizeDialog.fireOnOpen(
			HeaderSizeCtrl.newArg(Integer.valueOf(headerType), sameVal ? prev : null, selection));
	}

	public void openRenameSheetDialog(String originalSheetName) {
		if (_renameSheetDialog == null || _renameSheetDialog.isInvalidated())
			_renameSheetDialog = (Dialog) Executions.createComponents(Consts._RenameDialog_zul, mainWin, null);
		_renameSheetDialog.fireOnOpen(RenameSheetCtrl.newArg(originalSheetName));
	}

	public void openOpenFileDialog() {
		if (_openFileDialog == null || _openFileDialog.isInvalidated())
			_openFileDialog = (Dialog) Executions.createComponents(Consts._OpenFile_zul, mainWin, null);
		_openFileDialog.fireOnOpen(null);
	}

	public void openImportFileDialog() {
		if (_importFileDialog == null || _importFileDialog.isInvalidated())
			_importFileDialog = (Dialog) Executions.createComponents(Consts._ImportFile_zul, mainWin, null);
		_importFileDialog.fireOnOpen(null);
	}
	
	public void openExportPdfDialog(Rect selection) {
		if (!hasZssPdf()) {
			Messagebox.show("Please download Zss Pdf from ZK");
			return;
		}
		if (_exportToPdfDialog == null || _exportToPdfDialog.isInvalidated())
			_exportToPdfDialog = (Dialog) Executions.createComponents(Consts._ExportToPDF_zul, mainWin, Zssapps.newSpreadsheetArg(spreadsheet));
		_exportToPdfDialog.fireOnOpen(selection);
	}
		
	private static boolean hasZssPdf() {
		String val = Library.getProperty(KEY_PDF);
		if (val == null) {
			boolean hasZssPdf = verifyZssPdf();
			Library.setProperty(KEY_PDF, String.valueOf(hasZssPdf));
			return hasZssPdf;
		} else {
			return Boolean.valueOf(Library.getProperty(KEY_PDF));
		}
	}

	/**
	 * Verify whether has zss pdf export function or not
	 */
	private static boolean verifyZssPdf() {
		try {
			Class.forName("org.zkoss.zss.model.impl.pdf.PdfExporter");
		} catch (ClassNotFoundException ex) {
			return false;
		}
		return true;
	}

	//TODO: mimic openExportPdfDialog()
	@Override
	public void openExportHtmlDialog(Rect selection) {
		if (!hasZssHtml()) {
			Messagebox.show("Please download Zss Html from ZK");
			return;
		}
		if (_exportToHtmlDialog == null || _exportToHtmlDialog.isInvalidated())
			_exportToHtmlDialog = (Dialog) Executions.createComponents(Consts._ExportToHTML_zul, mainWin, Zssapps.newSpreadsheetArg(spreadsheet));
		_exportToHtmlDialog.fireOnOpen(null);
		
	}

	private static boolean hasZssHtml() {
		String val = Library.getProperty(KEY_HTML);
		if (val == null) {
			boolean hasZssHtml = verifyZssHtml();
			Library.setProperty(KEY_HTML, String.valueOf(hasZssHtml));
			return hasZssHtml;
		} else {
			return Boolean.valueOf(Library.getProperty(KEY_HTML));
		}
	}

	private static boolean verifyZssHtml() {
		try {
			Class.forName("org.zkoss.zss.model.impl.html.HtmlExporter");
		} catch (ClassNotFoundException ex) {
			return false;
		}
		return true;
	}
	
	public void onSheetSelect$spreadsheet() {
		getDesktopWorkbenchContext().fireWorkbookChanged();
	}
	
	private class MainActionHandler extends XActionHandler {
		
		MainActionHandler() {
			super(spreadsheet);
		}

		@Override
		public void doNewBook() {
			getDesktopWorkbenchContext().getWorkbookCtrl().newBook();
			getDesktopWorkbenchContext().fireWorkbookChanged();
		}

		@Override
		public void doSaveBook() {
			if (FileHelper.hasSavePermission() && spreadsheet.getXBook() != null) {
				DesktopWorkbenchContext workbench = getDesktopWorkbenchContext();
				if (workbench.getWorkbookCtrl().hasFileExtentionName()) {
					workbench.getWorkbookCtrl().save();
					workbench.fireWorkbookSaved();
				} else
					workbench.getWorkbenchCtrl().openSaveFileDialog();	
			}
		}

		@Override
		public void doCloseBook() {
			super.doCloseBook();
			
			fileMenu.setSaveFileDisabled(true);
			fileMenu.setSaveFileAndCloseDisabled(true);
			fileMenu.setDeleteFileDisabled(true);
			fileMenu.setExportPdfDisabled(true);
			fileMenu.setExportHtmlDisabled(true);
			fileMenu.setExportExcelDisabled(true);
			
			viewMenu.setFreezeUnFreezeDisabled(true);
		}

		@Override
		public void doExportPDF(Rect selection) {
			if (spreadsheet.getXBook() != null) {
				openExportPdfDialog(selection);	
			}
		}

		@Override
		public void doPasteSpecial(Rect selection) {
			Clipboard copyFrom = getClipboard();
			if (copyFrom == null) {
				Messagebox.show("Spreadsheet must has highlight area as paste source, please set spreadsheet's highlight area");
				return;
			}

			if (spreadsheet.getSelectedXSheet() != null) {
				spreadsheet.setSelection(selection);
				openPasteSpecialDialog();	
			}
		}
		
		

		@Override
		public void doCtrlKey(KeyEvent event) {
			super.doCtrlKey(event);
			if (spreadsheet.getSelectedXSheet() != null) {
				switch (event.getKeyCode()) {
				case 'S':
					//TODO: check permission from WorkbookCtrl
					if (FileHelper.hasSavePermission()) {
						//TODO: refactor duplicate save logic
						DesktopWorkbenchContext workbench = getDesktopWorkbenchContext();
						if (workbench.getWorkbookCtrl().hasFileExtentionName()) {
							workbench.getWorkbookCtrl().save();
							workbench.fireWorkbookSaved();
						} else
							workbench.getWorkbenchCtrl().openSaveFileDialog();
					}
					break;
				case 'O':
					openOpenFileDialog();
					break;
				}
			}
		}

		@Override
		public void doFontSize(int fontSize, Rect selection) {
			super.doFontSize(fontSize, selection);
			
			//For performance reason, markout this behavior. (when select 2 columns and set font size)
			//TODO: improve client side performance for this behavior
			//setProperRowHeightByFontSize(spreadsheet.getSelectedSheet(), selection, fontSize);
		}
		
		private void setProperRowHeightByFontSize(XSheet sheet, Rect rect, int size) {	
			int tRow = rect.getTop();
			int bRow = rect.getBottom();
			int col = rect.getLeft();
			
			for (int i = tRow; i <= bRow; i++) {
				//Note. add extra padding height: 4
				if ((size + 4) > (Utils.pxToPoint(Utils.twipToPx(BookHelper.getRowHeight(sheet, i))))) {
					XRanges.range(sheet, i, col).setRowHeight(size + 4);
				}
			}
		}

		@Override
		public void doCustomSort(Rect selection) {
			if (spreadsheet.getSelectedXSheet() != null) {
				openCustomSortDialog(selection);
			}
		}

		@Override
		public void doHyperlink(Rect selection) {
			if (spreadsheet.getSelectedXSheet() != null) {
				openHyperlinkDialog(selection);
			}
		}

		@Override
		public void doFormatCell(Rect selection) {
			if (spreadsheet.getSelectedXSheet() != null) {
				openFormatNumberDialog(selection);	
			}
		}

		@Override
		public void doColumnWidth(Rect selection) {
			if (spreadsheet.getSelectedXSheet() != null) {
				openModifyHeaderSizeDialog(WorkbookCtrl.HEADER_TYPE_COLUMN, selection);	
			}
		}

		@Override
		public void doRowHeight(Rect selection) {
			if (spreadsheet.getSelectedXSheet() != null) {
				openModifyHeaderSizeDialog(WorkbookCtrl.HEADER_TYPE_ROW, selection);	
			}
		}

		@Override
		public void doInsertFunction(Rect selection) {
			openInsertFormulaDialog(selection);
		}

		@Override
		public void toggleActionOnSheetSelected() {
			super.toggleActionOnSheetSelected();
			
//			boolean savePermission = FileHelper.hasSavePermission();
			//save button will enable onContentChange
			getSpreadsheet().setActionDisabled(true, UserAction.SAVE_BOOK);
		}
	}
}