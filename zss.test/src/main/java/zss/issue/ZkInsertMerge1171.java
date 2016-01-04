package zss.issue;

import java.io.File;

import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.EventQueue;
import org.zkoss.zk.ui.event.EventQueues;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.util.Clients;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.CellRef;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.InsertCopyOrigin;
import org.zkoss.zss.api.Range.InsertShift;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.CellMouseEvent;
import org.zkoss.zul.Vlayout;

public class ZkInsertMerge1171 extends SelectorComposer<Vlayout> {

	public final static String EXCEL_FILE = Executions.getCurrent().getDesktop().getWebApp().getRealPath("/issue3/book/1171-ff-slow-merge.xlsx");
	
	public final static String SYMBOL_COLLAPSED = "\u25C4\u25BA";
	public final static String SYMBOL_EXPANDED = "\u25BA\u25C4";

		private Spreadsheet spreadsheet;
		private Sheet selectedSheet;
		private int left, top, right, bottom;
		
		private EventQueue<Event> queue = EventQueues.lookup("heavyOperation");
		
		@Override
	    public void doAfterCompose(Vlayout comp) throws Exception {
			super.doAfterCompose(comp);
			spreadsheet = (Spreadsheet) comp.getFellowIfAny("sSpreadsheet");
			Book book = Importers.getImporter().imports(new File(EXCEL_FILE), "test");
			spreadsheet.setBook(book);
			
			selectedSheet = spreadsheet.getSelectedSheet();
	        
			//Emulating spreadsheet formatting and data population
	        spreadsheet.getBook().getInternalBook().createSheet("MAIN");
	        spreadsheet.setSelectedSheet("MAIN");
	        Sheet newSheet = spreadsheet.getSelectedSheet();
	        
			Range srcRange = Ranges.range(selectedSheet, 0, 0, selectedSheet.getLastRow(), selectedSheet.getLastColumn(10));
			Range destRange = Ranges.range(newSheet, 0, 0, selectedSheet.getLastRow(), selectedSheet.getLastColumn(10));
			srcRange.pasteSpecial(destRange, Range.PasteType.ALL, Range.PasteOperation.ADD, false, false);
			
			for(int i=destRange.getColumn(); i<destRange.getLastColumn()+1; i++){
				Ranges.range(newSheet, 0, i, 0, i).toColumnRange().setColumnWidth(selectedSheet.getColumnWidth(i));
			}
			
			for(int i=destRange.getRow(); i<destRange.getLastRow()+1; i++){
				Ranges.range(newSheet, 0, i, 0, i).toRowRange().setRowHeight(selectedSheet.getRowHeight(i));
			}
			Ranges.range(newSheet,0,0).toColumnRange().setHidden(true);
			selectedSheet = newSheet;
			
			spreadsheet.setMaxVisibleColumns(selectedSheet.getLastColumn(13)+20);
			spreadsheet.setMaxVisibleRows(selectedSheet.getLastRow());
			
			spreadsheet.setMaxRenderedCellSize(1500);
	        this.spreadsheet.setPreloadColumnSize(24);
	        this.spreadsheet.setPreloadRowSize(100);
	        Range r = Ranges.range(this.selectedSheet);
		    r.protectSheet("test", true, true, false, false, false, false, false, false, false, false, true, true, false, false, false);
		    
		    queue.subscribe(new EventListener<Event>() {
				
				@Override
				public void onEvent(Event e) throws Exception {
						performHeavyOperation((CellMouseEvent)e);
						//queue.publish(new Event("complete"));
				}
			}, new EventListener<Event>() {
				@Override
				public void onEvent(Event e) throws Exception {					
					spreadsheet.setMaxVisibleColumns(selectedSheet.getLastColumn(13)+20);
					spreadsheet.setMaxVisibleRows(selectedSheet.getLastRow());
					Range range = Ranges.range(selectedSheet, top, left, bottom, right);
					range.notifyChange();
//					Ranges.range(spreadsheet.getSelectedSheet()).notifyChange();
					Clients.clearBusy();
				}
			});
		}
		
		@Listen("onCellClick = #sSpreadsheet")
		public void onCellClick(CellMouseEvent event){
			if (event.getColumn() > 46 && event.getRow() == 10){ //header row
				queue.publish(event);
				Clients.showBusy("inserting");
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
				
				left = event.getColumn();
				top = 0;
				right = left + 9;
				bottom = 10;

				range = Ranges.range(selectedSheet, 0, left + 1, 0, left + 9).toColumnRange();
				range.setAutoRefresh(false);
				CellOperationUtil.insert(range, InsertShift.RIGHT, InsertCopyOrigin.FORMAT_LEFT_ABOVE);
				
				range = Ranges.range(selectedSheet, top, left, bottom, right);
				range.setAutoRefresh(false);
				CellOperationUtil.merge(range, true); //or range.merge(true);
								
				sheetRange.protectSheet("test", true, true, false, false, false, false, false, false, false, false, true, true, false, false, false);
				
				cellRange.setCellValue(SYMBOL_EXPANDED + " " + cellRange.getCellValue().toString().substring(3));
			}
		}
	
}
