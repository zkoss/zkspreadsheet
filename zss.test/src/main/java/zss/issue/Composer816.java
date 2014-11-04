/**
 * This is about Composer816
 * @author mqiu
 * @date 10/22/2014
 * 
 * Description: a demo to reseach the performance issue
 * and found out it related to r.protectSheet(...) 
 * with this setting it costs 45s while without only 3s
 */

package zss.issue;

import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Window;

public class Composer816 extends SelectorComposer<Window> {

	private static final long serialVersionUID = -7360829196117880724L;
	
	private static final int START_ROW = 0;
	private static final int TITLE_END_ROW = 1;
	private static final int CONTENT_ROW = 2;
	private static final int START_COLUMN = 0;
	private static final int LAST_COLUMN = START_COLUMN + 17;
	
	private static final int CONTENT_NUM = 300;
	
	private int lastRow = CONTENT_ROW;
	private int lastColumn = Composer816.LAST_COLUMN;
	
	
	@Wire
	Spreadsheet spreadsheet;
	
	Sheet srcSheet;
	Sheet sheet;
	
	public void doAfterCompose(Window comp) throws Exception {
        super.doAfterCompose(comp);
        srcSheet = spreadsheet.getSelectedSheet();
        spreadsheet.getBook().getInternalBook().createSheet("Demo");
        sheet = spreadsheet.getBook().getSheet("Demo");
        spreadsheet.setSelectedSheet("Demo");
        
        CellRegion titleRegion = new CellRegion(START_ROW, START_COLUMN, TITLE_END_ROW, LAST_COLUMN);
        copyTitle(titleRegion, titleRegion);
        
        for(int i=0; i < Composer816.CONTENT_NUM; i++){
        	copyRow(CONTENT_ROW, lastRow+1);
        }
        
        spreadsheet.getBook().getInternalBook().deleteSheet(srcSheet.getInternalSheet());
        
        Range r = Ranges.range(this.sheet);
        
        //this setting may led to a performance issue
		r.protectSheet("", false, true, false, false, false, false, false, false, false, false, false, false, false, false, false);
		
        spreadsheet.setMaxVisibleColumns(lastColumn);
        spreadsheet.setMaxVisibleRows(lastRow);
    }

	private void copyTitle(CellRegion srcRegion, CellRegion destRegion){
		copyRange(srcRegion, destRegion);
		copyRowHeight(srcRegion, destRegion);
		copyColumnWidth(srcRegion, destRegion);
	}
	
	private void copyRow(int srcRow, int destRow){
		CellRegion srcRegion = new CellRegion(srcRow, Composer816.START_COLUMN, srcRow, Composer816.LAST_COLUMN);
		CellRegion destRegion = new CellRegion(destRow, Composer816.START_COLUMN, destRow, Composer816.LAST_COLUMN);
		
		copyRange(srcRegion, destRegion);
		copyRowHeight(srcRegion, destRegion);
		
		lastRow=(lastRow<destRow?destRow:lastRow);
	}
	
	private void copyRange(CellRegion srcRegion, CellRegion destRegion){
		Range srcRange = Ranges.range(srcSheet, srcRegion.getReferenceString());
		Range destRange = Ranges.range(sheet, destRegion.getReferenceString());
		srcRange.pasteSpecial(destRange, Range.PasteType.ALL, Range.PasteOperation.ADD, false, false);
	}
	
	private void copyColumnWidth(CellRegion srcRegion, CellRegion destRegion){
		SSheet ssrc = srcSheet.getInternalSheet();
		SSheet sdest = sheet.getInternalSheet();
		for(int i=destRegion.column; i<destRegion.lastColumn+1; i++){
		    //I can't setCustomWidth true by range, so here I use SSheet, it means this function only work when init the sheet
//			Range destRange = Ranges.range(dest, i, destRegion.column, i, destRegion.lastColumn);
			int srcColumn = srcRegion.column + i - destRegion.column;
//			destRange.setColumnWidth(src.getColumnWidth(srcColumn));
			sdest.getColumn(i).setCustomWidth(true);
			sdest.getColumn(i).setWidth(ssrc.getColumn(srcColumn).getWidth());
		}
	}
	
	//It's wired that we can use range to setCustonRowHeight but can't setColumnWidth
	private void copyRowHeight(CellRegion srcRegion, CellRegion destRegion){
		for(int i=destRegion.row; i<destRegion.lastRow+1; i++){
			Range destRange = Ranges.range(sheet, i, destRegion.column, i, destRegion.lastColumn);
			int srcRow = srcRegion.row + i - destRegion.row;
			destRange.setRowHeight(srcSheet.getRowHeight(srcRow), true);			
		}
	}
}
