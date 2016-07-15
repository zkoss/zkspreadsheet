package zss.issue;

import java.io.File;
import java.io.InputStream;

import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.CellRef;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.Range.InsertCopyOrigin;
import org.zkoss.zss.api.Range.InsertShift;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.CellMouseEvent;
import org.zkoss.zul.Vlayout;

public class ZKInsertMerge1115 extends SelectorComposer<Vlayout> {
	
	private final String EXCELFILE = "/issue3/book/1115-zkinsertmerge.xlsx";
    public final static String SYMBOL_COLLAPSED = "\u25C4\u25BA ";
    public final static String SYMBOL_EXPANDED  = "\u25BA\u25C4 ";
    
    private Spreadsheet spreadsheet;
	private Sheet selectedSheet;
	
	@Override
    public void doAfterCompose(Vlayout comp) throws Exception {
		super.doAfterCompose(comp);
		spreadsheet = (Spreadsheet) comp.getFellowIfAny("spreadsheet");
		InputStream is = Executions.getCurrent().getDesktop().getWebApp().getResourceAsStream(EXCELFILE);
		Book book = Importers.getImporter().imports(is, "test");
		spreadsheet.setBook(book);
		selectedSheet = spreadsheet.getSelectedSheet();
        
		//Emulating spreadsheet formatting and data population
        book.getInternalBook().createSheet("MAIN");
        spreadsheet.setSelectedSheet("MAIN");
        Sheet newSheet = spreadsheet.getSelectedSheet();
        
		Range srcRange = Ranges.range(selectedSheet, 0, 0, selectedSheet.getLastRow(), selectedSheet.getLastColumn(10));
		Range destRange = Ranges.range(newSheet, 0, 0, selectedSheet.getLastRow(), selectedSheet.getLastColumn(10));
		srcRange.pasteSpecial(destRange, Range.PasteType.ALL, Range.PasteOperation.ADD, false, false);
		
		for(int i=destRange.getColumn(); i<destRange.getLastColumn()+1; i++){
			Ranges.range(newSheet, 0, i, 0, i).toColumnRange().setColumnWidth(selectedSheet.getColumnWidth(i));
		}
		SSheet ssheet = selectedSheet.getInternalSheet();
		for(int i=destRange.getRow(); i<destRange.getLastRow()+1; i++){
			final boolean isCustom = ssheet.getRow(i).isCustomHeight();
			Ranges.range(newSheet, i, 0, i, 0).toRowRange().setRowHeight(selectedSheet.getRowHeight(i), isCustom);
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
	}
	
	@Listen("onCellClick = #spreadsheet")
	public void onCellClick(CellMouseEvent event){
		if (event.getColumn() > 46 && event.getRow() == 10){ //header row, AV ~
			Range cellRange = Ranges.range(this.selectedSheet, event.getRow(), event.getColumn());
			if (cellRange.getCellValue() != null && cellRange.getCellValue().toString().startsWith(SYMBOL_COLLAPSED)){
				Ranges.range(this.selectedSheet).unprotectSheet("test");
				final Range insertRange = Ranges.range(selectedSheet, 0, event.getColumn() + 1, 0, event.getColumn() + 9).toColumnRange();
				insertRange.setAutoRefresh(false);
				CellOperationUtil.insert(insertRange, InsertShift.RIGHT, InsertCopyOrigin.FORMAT_LEFT_ABOVE);			
				Ranges.range(this.selectedSheet).protectSheet("test", true, true, false, false, false, false, false, false, false, false, true, true, false, false, false);
				
				cellRange.setAutoRefresh(false);
				cellRange.setCellValue(SYMBOL_EXPANDED + " " + cellRange.getCellValue().toString().substring(3));
				spreadsheet.setMaxVisibleColumns(selectedSheet.getLastColumn(13)+20);
				spreadsheet.setMaxVisibleRows(selectedSheet.getLastRow());
				
				Ranges.range(this.selectedSheet).notifyChange();
			}
		}
	}
}
