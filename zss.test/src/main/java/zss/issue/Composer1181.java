package zss.issue;

import java.io.File;

import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zk.ui.util.Clients;
import org.zkoss.zss.api.AreaRef;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.CellRef;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.CellAttribute;
import org.zkoss.zss.api.Range.DeleteShift;
import org.zkoss.zss.api.Range.InsertCopyOrigin;
import org.zkoss.zss.api.Range.InsertShift;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.CellStyle.Alignment;
import org.zkoss.zss.api.model.CellStyle.VerticalAlignment;
import org.zkoss.zss.api.model.Hyperlink.HyperlinkType;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.CellMouseEvent;
import org.zkoss.zss.ui.event.StopEditingEvent;
import org.zkoss.zss.ui.impl.ActiveRangeHelper;
import org.zkoss.zss.ui.sys.FreezeInfoLoader;
import org.zkoss.zss.ui.sys.SpreadsheetCtrl;
import org.zkoss.zul.Vlayout;

public class Composer1181 extends SelectorComposer<Vlayout> {

	public final static String EXCEL_FILE = Executions.getCurrent().getDesktop().getWebApp().getRealPath("/issue3/book/1181-zkinsertmerge.xlsx");
	
	public final static String SYMBOL_COLLAPSED = "\u25C4\u25BA";
	public final static String SYMBOL_EXPANDED = "\u25BA\u25C4";
	
	public final static int ROW_HEADER_1 = 10;
	public final static int ROW_HEADER_2 = ROW_HEADER_1 + 1;

	@Wire("#sSpreadsheet")
	private Spreadsheet spreadsheet;
	private Sheet selectedSheet;

	@Override
	public void doAfterCompose(Vlayout comp) throws Exception {
		super.doAfterCompose(comp);
		Book book = Importers.getImporter().imports(new File(EXCEL_FILE), "test");

		initializeData(book);
		initializeSpreadsheet(selectedSheet);
	}

	private void initializeSpreadsheet(Sheet selectedSheet) {
		spreadsheet.setMaxVisibleColumns(selectedSheet.getLastColumn(13)+20);
		spreadsheet.setMaxVisibleRows(selectedSheet.getLastRow());
	}
	
	private void initializeData(Book book){
		spreadsheet.setBook(book);
		selectedSheet = spreadsheet.getSelectedSheet();
		//Emulating spreadsheet formatting and data population
        spreadsheet.getBook().getInternalBook().createSheet("MAIN");
        spreadsheet.setSelectedSheet("MAIN");
        Sheet newSheet = spreadsheet.getSelectedSheet();
		newSheet.getInternalSheet().getViewInfo().setDisplayGridlines(false);
		
		Range srcRange = getRange(selectedSheet, 0, 0, selectedSheet.getLastRow(), selectedSheet.getLastColumn(10));
		Range destRange = getRange(newSheet, 0, 0, selectedSheet.getLastRow(), selectedSheet.getLastColumn(10));
		srcRange.pasteSpecial(destRange, Range.PasteType.ALL, Range.PasteOperation.ADD, false, false);
		
		for(int i=destRange.getColumn(); i<destRange.getLastColumn()+1; i++){
			Ranges.range(newSheet, 0, i, 0, i).toColumnRange().setColumnWidth(selectedSheet.getColumnWidth(i));
		}
		
		for(int i=destRange.getRow(); i<destRange.getLastRow()+1; i++){
			Ranges.range(newSheet, 0, i, 0, i).toRowRange().setRowHeight(selectedSheet.getRowHeight(i));
		}
		
		int startRow = ROW_HEADER_2 + 1;
		for(int row = startRow; row < destRange.getLastRow()+1; row++){
			Range range = Ranges.range(newSheet, row, 2);
			if (range.getCellValue() != null){
				range.setCellHyperlink(HyperlinkType.FILE, "#", (String)range.getCellValue());
			} else {
				Range rangeMerge = Ranges.range(newSheet, startRow, 1, row - 1, 1);
				rangeMerge.merge(false);
				rangeMerge = Ranges.range(newSheet, startRow, 2, row - 1, 2);
				rangeMerge.merge(false);
				CellOperationUtil.applyVerticalAlignment(rangeMerge, VerticalAlignment.CENTER);
				CellOperationUtil.applyAlignment(rangeMerge, Alignment.CENTER);
				startRow = row + 1;
			}
			
		}
		
		Ranges.range(newSheet,0,0).toColumnRange().setHidden(true);
		selectedSheet = newSheet;
		
        Range r = Ranges.range(this.selectedSheet);
	    r.protectSheet("test", true, true, false, false, false, false, false, false, false, false, true, true, false, false, false);		
	}

	@Listen("onCellClick = #sSpreadsheet")
	public void onCellClick(CellMouseEvent event){
		System.out.println("onCellClick:" +  new CellRef(event.getRow(), event.getColumn()).asString());

		if (event.getColumn() > 46 && event.getRow() == 10){ //header row
			Range range = getRange(selectedSheet, event.getRow(), event.getColumn());
			if (range.getCellValue() == null){ return; }
			String cellValue = range.getCellValue().toString();
			if (cellValue.startsWith(SYMBOL_COLLAPSED)){
				Clients.showBusy("inserting");
				Events.echoEvent("onInsert", spreadsheet, event);
			} else if (cellValue.startsWith(SYMBOL_EXPANDED)){
				Clients.showBusy("Delete column");
				Events.echoEvent("onDeleteColumn", spreadsheet, event);
			}
		}
	}

	public void performHeavyOperation(CellMouseEvent event){
		Range cellRange = Ranges.range(this.selectedSheet, event.getRow(), event.getColumn());
		cellRange.setAutoRefresh(false);
		if (cellRange.getCellValue() != null && cellRange.getCellValue().toString().startsWith(SYMBOL_COLLAPSED)){
			Range sheetRange = Ranges.range(this.selectedSheet);
			sheetRange.setAutoRefresh(false);
			sheetRange.unprotectSheet("test");

			Range range = null;

			range = Ranges.range(selectedSheet, 0, event.getColumn() + 1, 0, event.getColumn() + 9).toColumnRange();
//			range = Ranges.range(selectedSheet, 0, event.getColumn() + 1, selectedSheet.getLastRow(), event.getColumn() + 9);
			range.setAutoRefresh(false);
			CellOperationUtil.insert(range, InsertShift.RIGHT, InsertCopyOrigin.FORMAT_LEFT_ABOVE);

			range = Ranges.range(selectedSheet, 0, event.getColumn(), 10, event.getColumn() + 9);
			range.setAutoRefresh(false);
			CellOperationUtil.merge(range, true);

			sheetRange.protectSheet("test", true, true, false, false, false, false, false, false, false, false, true, true, false, false, false);

			cellRange.setCellValue(SYMBOL_EXPANDED + " " + cellRange.getCellValue().toString().substring(3));
		}
	}
	
	@Listen("onInsert = #sSpreadsheet")
	public void merge(Event e){
		CellMouseEvent event = (CellMouseEvent)e.getData();
		performHeavyOperation(event);
		// option1
//		notifyAffectedRange(event);
		// option2
//		spreadsheet.notifyLoadedAreaChange();
		// option3
		spreadsheet.notifyVisibleAreaChange();
		
		spreadsheet.setMaxVisibleColumns(selectedSheet.getLastColumn(13)+20);
		spreadsheet.setMaxVisibleRows(selectedSheet.getLastRow());
		Clients.clearBusy();
	}

	private void notifyAffectedRange(CellMouseEvent event) {
		Ranges.range(spreadsheet.getSelectedSheet(), 0, event.getColumn(), spreadsheet.getSelectedSheet().getLastRow(),
				spreadsheet.getSelectedSheet().getLastColumn(13)+20).notifyChange();
		System.out.println(new AreaRef( 0, event.getColumn(), spreadsheet.getSelectedSheet().getLastRow(),
				spreadsheet.getSelectedSheet().getLastColumn(13)+20).asString());
	}

	private void notifyLoadedRange() {
		SpreadsheetCtrl ssctrl = (SpreadsheetCtrl) spreadsheet.getExtraCtrl();
		AreaRef area = ssctrl.getLoadedArea();
		FreezeInfoLoader freezeInfo = ssctrl.getFreezeInfoLoader();
		
		Ranges.range(spreadsheet.getSelectedSheet(), ssctrl.getLoadedArea()).notifyChange();
		System.out.println(((SpreadsheetCtrl) spreadsheet.getExtraCtrl()).getLoadedArea().asString());
	}
	
	@Listen("onDeleteColumn = #sSpreadsheet")
	public void onDeleteColumn(Event e){
		CellMouseEvent event = (CellMouseEvent)e.getData();
		Range sheetRange = Ranges.range(this.selectedSheet);
		sheetRange.setAutoRefresh(false);
		sheetRange.unprotectSheet("test");

		CellOperationUtil.delete(this.getColumnRange(selectedSheet, 0, event.getColumn() + 1, 0, event.getColumn() + 9), 
								 DeleteShift.LEFT);
		
		Range cellRange = this.getRange(this.selectedSheet, event.getRow(), event.getColumn());
		cellRange.setCellValue(SYMBOL_COLLAPSED + " " + cellRange.getCellValue().toString().substring(3));
		
		sheetRange.protectSheet("test", true, true, false, false, false, false, false, false, false, false, true, true, false, false, false);
		
		Ranges.range(selectedSheet, 0, event.getColumn(), selectedSheet.getLastRow(), selectedSheet.getLastColumn(13)+20).notifyChange();
//		Ranges.range(spreadsheet.getSelectedSheet(), ((SpreadsheetCtrl) spreadsheet.getExtraCtrl()).getLoadedArea()).notifyChange();
		spreadsheet.setMaxVisibleColumns(selectedSheet.getLastColumn(13)+20);
		spreadsheet.setMaxVisibleRows(selectedSheet.getLastRow());
		Clients.clearBusy();
	}
	
	@Listen("onStopEditing = #sSpreadsheet")
	public void onStopEdit(StopEditingEvent event){
		System.out.println("onCellClick:" +  new CellRef(event.getRow(), event.getColumn()).asString());

		if (event.getColumn() > 46 && event.getRow() == 12){ //first row		
			Clients.showBusy("Update total");
			Events.echoEvent("onUpdateTotal", spreadsheet, event);
		}
	}
	
	@Listen("onUpdateTotal = #sSpreadsheet")
	public void onUpdateTotal(Event e){
		StopEditingEvent event = (StopEditingEvent)e.getData();
		String editingValue = event.getEditingValue().toString();
		int row = event.getRow();
		
		for (int i = 36; i < 47; i++){
			getRange(selectedSheet, row, i).setCellValue(editingValue);
		}
		
		int lastCol = selectedSheet.getLastColumn(ROW_HEADER_1);
		for (int i = 47; i < lastCol; i++){
			getRange(selectedSheet, 17, i).setCellValue(editingValue);
		}
		
		Ranges.range(selectedSheet, 0, 0, selectedSheet.getLastRow(), selectedSheet.getLastColumn(ROW_HEADER_1)).notifyChange();
		Clients.clearBusy();
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
