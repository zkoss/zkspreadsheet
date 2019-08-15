/**
 * 
 */
package zss.issue;

import java.io.File;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.PostImport;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.SheetVisible;
import org.zkoss.zss.api.Range.SortDataOption;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.ui.AuxAction;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Label;

@SuppressWarnings("serial")
public class Composer1283PopulatePerformance extends SelectorComposer<Component> {
    private final static Importer FILE_IMPORTER = Importers.getImporter();
    
    @Wire
    private Spreadsheet ss;
    @Wire
    private Label metaLabel;
    private Book book;
    private Sheet sheet;

    @Listen("onClick = #postImport")
    public void onPostImport(Event event) throws Exception {
        Clock clock = new Clock();
        clock.start();
        File template = new File(WebApps.getCurrent().getRealPath("/book/blank.xlsx"));
        String templateDocName = "input";
//        this.book = FILE_IMPORTER.imports(template, templateDocName);
//        this.postImport(this.book);
        this.book = FILE_IMPORTER.imports(template, templateDocName, new PostImport() {
        	public void process(Book book) {
        		Composer1283PopulatePerformance.this.postImport(book);
        	}
        });
        clock.stop();
        metaLabel.setValue("consumedtime: " + clock.elaspsedTimeString());
    }
    @Listen("onClick = #usual")
    public void onUsual(Event event) throws Exception {
        Clock clock = new Clock();
        clock.start();
        File template = new File(WebApps.getCurrent().getRealPath("/book/blank.xlsx"));
        String templateDocName = "input";
        this.book = FILE_IMPORTER.imports(template, templateDocName);
        this.postImport(this.book);
//        this.book = FILE_IMPORTER.imports(template, templateDocName, new PostImport() {
//        	public void process(Book book) {
//        		Composer1283PopulatePerformance.this.postImport(book);
//        	}
//        });
        clock.stop();
        metaLabel.setValue("consumedtime: " + clock.elaspsedTimeString());
    }
    

//    @Override
//    public void doAfterCompose(Component comp) throws Exception {
//        super.doAfterCompose(comp);
//        File template = new File(WebApps.getCurrent().getRealPath("/book/blank.xlsx"));
//        String templateDocName = "input";
////        this.book = FILE_IMPORTER.imports(template, templateDocName);
////        this.postImport(this.book);
//        this.book = FILE_IMPORTER.imports(template, templateDocName, new PostImport() {
//        	public void process(Book book) {
//        		Composer1283PopulatePerformance.this.postImport(book);
//        	}
//        });
//    }
    private void postImport(Book book0) {
//        hideSheet(book.getSheet("input"));
//        String sheetName = "worksheet";
//        getRange(this.book.getSheetAt(0)).createSheet(sheetName);
    	this.book = book0;
        this.sheet = this.book.getSheetAt(0);
        
        
        //FIXME: @Hawk, In order to make demo simple, here I just use simple formula to reproduce the performance issue. In our product, we have very complex formula. 
        //The Characteristics of formula performance issues:
        //1. The performance becomes bad if enable formula.
        //2. The performance gets more worse when rows increasing. The code is processing row by row. When processing the latter rows, the time of processing it is signaficantly increased. 
        //I attach the example in the ticket by image. It seems that it costs much memory with formula.
        //You can change isEnabledFormula to test with formula and without formula cases.
        
//        fillFormulaAtInternalSheet(sheet); workaround
        boolean isEnabledFormula = true;
        int lastRow = 8000; //14000;
        int lastColumn = 256;
        Clock clock = new Clock();
        clock.start();
        Clock recordClockPer1000 = new Clock();
        recordClockPer1000.start();
        for(int row = 0; row < lastRow; row++) {
            if( row > 0 && (row+1)%1000 == 0) {
                recordClockPer1000.stop();
                System.out.println("consumedtime for " + (row - 999) + "-" + (row+1) + " rows : " + recordClockPer1000.elaspsedTimeString());
                recordClockPer1000.start();
            }
            String formula = "";
            for(int column = 0; column <  lastColumn; column++ ) {
                if(row > 0 && row%3 == 0) {
                    formula = "=sum(" + Ranges.getCellRefString(row-1, column) + "," + Ranges.getCellRefString(row-2, column) + ")" ;
                }
                if(row > 0 && row%3 == 0 && isEnabledFormula) {
                    getRange(this.sheet, row, column).setCellValue(formula); 
                }
                else {
                    getRange(this.sheet, row, column).setCellValue(row+column);
                }
            }
        }
        clock.stop();
        System.out.println("consumedtime: " + clock.elaspsedTimeString());
        //to disable actions in sheet context menu
        this.ss.disableUserAction(AuxAction.ADD_SHEET, true);
        this.ss.disableUserAction(AuxAction.COPY_SHEET, true);
        this.ss.disableUserAction(AuxAction.DELETE_SHEET, true);
        this.ss.disableUserAction(AuxAction.RENAME_SHEET, true);
        this.ss.disableUserAction(AuxAction.HIDE_SHEET, true);
        this.ss.disableUserAction(AuxAction.UNHIDE_SHEET, true);
        this.ss.disableUserAction(AuxAction.PROTECT_SHEET, true);
        this.ss.disableUserAction(AuxAction.MOVE_SHEET_LEFT, true);
        this.ss.disableUserAction(AuxAction.MOVE_SHEET_RIGHT, true);
        
        this.ss.setMaxVisibleColumns(lastColumn+1);
        this.ss.setMaxVisibleRows(lastRow+1);
        this.ss.setBook(this.book);
//        getRange(book.getSheet("input")).deleteSheet();
        getRange(this.sheet).notifyChange();
    }
    
   
    //workaround
    private void fillFormulaAtInternalSheet(Sheet sheet) {
    	SSheet internalSheet = sheet.getInternalSheet();
    	 int lastRow = 14000;
         int lastColumn = 256;
         Clock clock = new Clock();
         clock.start();
         Clock recordClockPer1000 = new Clock();
         recordClockPer1000.start();
         
         for(int row = 0; row < lastRow; row++) { 
             if( row > 0 && (row+1)%1000 == 0) {
                 recordClockPer1000.stop();
                 System.out.println("consumedtime for " + (row - 999) + "-" + (row+1) + " rows : " + recordClockPer1000.elaspsedTimeString());
                 recordClockPer1000.start();
             }
        	 for(int column = 0; column <  lastColumn; column++ ) {
        		 SCell cell = internalSheet.getCell(row, column);
        		 if(row > 0 && row%3 == 0) {
        			 String formula = "sum(" + Ranges.getCellRefString(row-1, column) + "," + Ranges.getCellRefString(row-2, column) + ")" ;
        			 cell.setFormulaValue(formula);
        		 }else{
        			 cell.setValue(row+column);
        		 }
        	 }
         }
         clock.stop();
         System.out.println("consumedtime: " + clock.elaspsedTimeString());
		
	}



	private  static Range getRange(Sheet sheet, int row, int col){
        Range range = Ranges.range(sheet, row, col);
        range.setAutoRefresh(false);
        return range;
    }


    private static Range getRange(Sheet sheet){
        Range range = Ranges.range(sheet);
        range.setAutoRefresh(false);
        return range;
    }
    
    private  static void hideSheet(Sheet sheet){
        getRange(sheet).setSheetVisible(SheetVisible.VERY_HIDDEN);
    }
}
