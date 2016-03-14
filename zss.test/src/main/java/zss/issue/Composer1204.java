package zss.issue;

import java.io.File;
import java.io.IOException;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zk.ui.util.Clients;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Vlayout;

public class Composer1204 extends SelectorComposer<Vlayout>{

	public final static String EXCEL_FILE = Executions.getCurrent().getDesktop().getWebApp().getRealPath("/issue3/book/1204-slow-merge.xlsx"); 
			//"/WEB-INF/issue/t2963.xlsx";
	
	public final static String SYMBOL_COLLAPSED = "\u25C4\u25BA";
	public final static String SYMBOL_EXPANDED = "\u25BA\u25C4";
	
	public final static int ROW_HEADER_1 = 10;
	public final static int ROW_HEADER_2 = ROW_HEADER_1 + 1;

	@Wire("#meroot")
	private Component meroot;
	@Wire("#sSpreadsheet")
	private Spreadsheet spreadsheet;
	
	private Book bookCache;
	
	private Sheet selectedSheet;
	
	@Override
	public void doAfterCompose(Vlayout comp) throws Exception {
		super.doAfterCompose(comp);
		onInit();
	}
		
	@Listen("onInit = #sSpreadsheet")
	public void onInit() throws IOException{
		if (bookCache == null){
			bookCache = Importers.getImporter().imports(new File(EXCEL_FILE), "test");
		}
		initializeData(bookCache);
		initializeSpreadsheet(selectedSheet);
	}
	
	public void initializeSpreadsheet(Sheet selectedSheet) {
		spreadsheet.setMaxVisibleColumns(selectedSheet.getLastColumn(13)+20);
		spreadsheet.setMaxVisibleRows(selectedSheet.getLastRow());
	}
	
	private void initializeData(Book book){
		spreadsheet.setBook(book);
		selectedSheet = spreadsheet.getSelectedSheet();
		//Emulating spreadsheet formatting and data population
		
		Sheet mainTempSheet = getRange(selectedSheet).createSheet("MAIN-TEMP");
		Sheet mainSheet = getRange(selectedSheet).createSheet("MAIN");

		Range srcRange = getRange(selectedSheet, 0, 0, selectedSheet.getLastRow(), selectedSheet.getLastColumn(10));
		Range destRange = getRange(mainTempSheet, 0, 0, selectedSheet.getLastRow(), selectedSheet.getLastColumn(10));
		srcRange.pasteSpecial(destRange, Range.PasteType.ALL, Range.PasteOperation.ADD, false, false);
		
		srcRange = getRange(mainTempSheet, 0, 0, mainTempSheet.getLastRow(), mainTempSheet.getLastColumn(10));
		destRange = getRange(mainSheet, 0, 0, mainTempSheet.getLastRow(), mainTempSheet.getLastColumn(10));
		srcRange.pasteSpecial(destRange, Range.PasteType.ALL, Range.PasteOperation.ADD, false, false);
		
		for(int i=destRange.getColumn(); i<destRange.getLastColumn()+1; i++){
			Ranges.range(mainSheet, 0, i, 0, i).toColumnRange().setColumnWidth(selectedSheet.getColumnWidth(i));
		}
		
		for(int i=destRange.getRow(); i<destRange.getLastRow()+1; i++){
			Ranges.range(mainSheet, 0, i, 0, i).toRowRange().setRowHeight(selectedSheet.getRowHeight(i));
		}
		
		Ranges.range(mainSheet,0,0).toColumnRange().setHidden(true);
		
		getRange(mainSheet).setSheetOrder(0);
		spreadsheet.setSelectedSheet("MAIN");
		selectedSheet = mainSheet;
		
       	getRange(this.selectedSheet).protectSheet("test", true, true, false, false, false, false, false, false, false, false, true, true, false, false, false);		
	}
		
	@Listen("onClick = #rebuildsheet")
	public void onClickDeleteSheet(Event e){
		Clients.showBusy(spreadsheet, "Rebuilding...");
		Events.echoEvent("onReBuildMainSheet", spreadsheet, null);
	}
	
	@Listen("onReBuildMainSheet = #sSpreadsheet")
	public void reBuildMainSheet(){
		Sheet mainTempSheet = spreadsheet.getBook().getSheet("MAIN-TEMP");
		Sheet mainSheet = spreadsheet.getBook().getSheet("MAIN");
		
		int mainIndex = spreadsheet.getBook().getSheetIndex(mainSheet);
		getRange(mainSheet).deleteSheet();
		
		mainSheet = getRange(mainTempSheet).createSheet("MAIN");
		getRange(mainSheet).setSheetOrder(mainIndex);
		//Emulating sheet formatting and data population
		int lastCol = mainTempSheet.getLastColumn(10);
		Range srcRange = getRange(mainTempSheet, 0, 0, 11, lastCol);
		Range destRange = getRange(mainSheet, 0, 0, 11, lastCol);
		srcRange.pasteSpecial(destRange, Range.PasteType.ALL, Range.PasteOperation.ADD, false, false);
		
		for (int row = 12; row < mainTempSheet.getLastRow(); row ++){
			srcRange = getRange(mainTempSheet, row, 0, row, lastCol);
			destRange = getRange(mainSheet, row, 0, row, lastCol);
			srcRange.pasteSpecial(destRange, Range.PasteType.ALL, Range.PasteOperation.ADD, false, false);
		}
		
		getRange(mainSheet).notifyChange();
		Clients.clearBusy(spreadsheet);
	}
	
	public Range getRange(Sheet sheet){
		Range range = Ranges.range(sheet);
		range.setAutoRefresh(false);
		return range;
	}
	
	public Range getRange(Sheet sheet, int row, int col){
		Range range = Ranges.range(sheet, row, col);
		range.setAutoRefresh(false);
		return range;
	}
	
	
	public Range getRange(Sheet sheet, int row, int col, int lastRow, int lastCol){
		Range range = Ranges.range(sheet, row, col, lastRow, lastCol);
		range.setAutoRefresh(false);
		return range;
	}
	
	public Range getColumnRange(Sheet sheet, int row, int col, int lastRow, int lastCol){
		Range range = Ranges.range(sheet, row, col, lastRow, lastCol).toColumnRange();
		range.setAutoRefresh(false);
		return range;
	}
	
	public Range getRowRange(Sheet sheet, int row, int col, int lastRow, int lastCol){
		Range range = Ranges.range(sheet, row, col, lastRow, lastCol).toRowRange();
		range.setAutoRefresh(false);
		return range;
	}

}
