/* XSSFSheetImpl.java

	Purpose:
		
	Description:
		
	History:
		Sep 3, 2010 11:07:42 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.sys.impl;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.SortedMap;
import java.util.TreeMap;

import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCols;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTComment;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPane;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetView;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetViews;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STPaneState;
import org.zkoss.lang.Classes;
import org.zkoss.lang.Library;
import org.zkoss.poi.POIXMLDocumentPart;
import org.zkoss.poi.POIXMLException;
import org.zkoss.poi.openxml4j.exceptions.InvalidFormatException;
import org.zkoss.poi.openxml4j.opc.PackagePart;
import org.zkoss.poi.openxml4j.opc.PackageRelationship;
import org.zkoss.poi.openxml4j.opc.PackageRelationshipCollection;
import org.zkoss.poi.ss.SpreadsheetVersion;
import org.zkoss.poi.ss.formula.FormulaParser;
import org.zkoss.poi.ss.formula.FormulaRenderer;
import org.zkoss.poi.ss.formula.FormulaType;
import org.zkoss.poi.ss.formula.PtgShifter;
import org.zkoss.poi.ss.formula.ptg.Ptg;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Chart;
import org.zkoss.poi.ss.usermodel.Drawing;
import org.zkoss.poi.ss.usermodel.Picture;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.poi.ss.util.CellReference;
import org.zkoss.poi.xssf.model.CommentsTable;
import org.zkoss.poi.xssf.usermodel.XSSFCell;
import org.zkoss.poi.xssf.usermodel.XSSFCellHelper;
import org.zkoss.poi.xssf.usermodel.XSSFComment;
import org.zkoss.poi.xssf.usermodel.XSSFDrawing;
import org.zkoss.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.zkoss.poi.xssf.usermodel.XSSFHyperlink;
import org.zkoss.poi.xssf.usermodel.XSSFName;
import org.zkoss.poi.xssf.usermodel.XSSFRelation;
import org.zkoss.poi.xssf.usermodel.XSSFRow;
import org.zkoss.poi.xssf.usermodel.XSSFRowHelper;
import org.zkoss.poi.xssf.usermodel.XSSFSheet;
import org.zkoss.poi.xssf.usermodel.XSSFWorkbook;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.XRange;
import org.zkoss.zss.model.sys.XSheet;

/**
 * Implementation of {@link XSheet} based on XSSFSheet.
 * @author henrichen
 *
 */
public class XSSFSheetImpl extends XSSFSheet implements SheetCtrl, XSheet {
	//--XSSFSheet--//
    public XSSFSheetImpl() {
        super();
    }

	/**
     * Creates an XSSFSheet representing the given package part and relationship.
     * Should only be called by XSSFWorkbook when reading in an exisiting file.
     *
     * @param part - The package part that holds xml data represenring this sheet.
     * @param rel - the relationship of the given package part in the underlying OPC package
     */
    public XSSFSheetImpl(PackagePart part, PackageRelationship rel) {
        super(part, rel);
    }

    @Override
    public int addMergedRegion(CellRangeAddress region)
    {
    	if (region != null) {
	    	addMerged(region);
    	}
        return super.addMergedRegion(region);
    }

    @Override
    public void removeMergedRegion(int index) {
    	final CellRangeAddress region = getMergedRegion(index);
    	if (region != null) {
        	deleteMerged(region);
    	}
    	super.removeMergedRegion(index);
    }

    //--Worksheet--//
	public XBook getBook() {
		return (XBook) getWorkbook();
	}

	@Override
	public List<Picture> getPictures() {
		DrawingManager dm = getDrawingManager();
		return new ArrayList<Picture>(dm.getPictures());
	}
	
	@Override
	public List<Chart> getCharts() {
		DrawingManager dm = getDrawingManager();
		return dm.getCharts();
	}
	
	//20100914, henrichen@zkoss.org: Shift rows only, don't handle formula
    /**
     * Shifts rows between startRow and endRow n number of rows.
     * If you use a negative number, it will shift rows up.
     * Code ensures that rows don't wrap around
     *
     * <p>
     * Additionally shifts merged regions that are completely defined in these
     * rows (ie. merged 2 cells on a row to be shifted).
     * <p>
     * @param startRow the row to start shifting
     * @param endRow the row to end shifting
     * @param n the number of rows to shift
     * @param copyRowHeight whether to copy the row height during the shift
     * @param resetOriginalRowHeight whether to set the original row's height to the default
     */
    public List<CellRangeAddress[]> shiftRowsOnly(int startRow, int endRow, int n, boolean copyRowHeight, boolean resetOriginalRowHeight,
    		boolean moveComments, boolean clearRest, int copyOrigin) {
    	
    	final int maxrow = SpreadsheetVersion.EXCEL2007.getLastRowIndex();
    	final int maxcol = SpreadsheetVersion.EXCEL2007.getLastColumnIndex();
    	
	    // ZSS-418: moving a comment must move both comment and shape positions
	    // Otherwise, Excel will fail to load such file
        if(moveComments) {
            shiftComments(startRow, maxrow, 0, maxcol, n, 0);
        }
    	
    	//prepare source format row
    	final int srcRownum = n <= 0 ? -1 : copyOrigin == XRange.FORMAT_RIGHTBELOW ? startRow : copyOrigin == XRange.FORMAT_LEFTABOVE ? startRow - 1 : -1;
    	final XSSFRow srcRow = srcRownum >= 0 ? getRow(srcRownum) : null;
    	final Map<Integer, Cell> srcCells = srcRow != null ? BookHelper.copyRowCells(srcRow, srcRow.getFirstCellNum(), srcRow.getLastCellNum()) : null;
    	final short srcHeight = srcRow != null ? srcRow.getHeight() : -1;
//    	final XSSFCellStyle srcStyle = srcRow != null ? srcRow.getRowStyle() : null;
    	
        final List<CellRangeAddress[]> shiftedRanges = BookHelper.shiftMergedRegion(this, startRow, 0, maxrow, maxcol, n, false);
    	
    	//shift the rows (actually change the row number only)
        List<Row> rowsToRemove = new ArrayList<Row>(); // ZSS-419: queue row for removing later, remove can't be perform in this loop 
        for (Iterator<Row> it = rowIterator() ; it.hasNext() ; ) {
            XSSFRow row = (XSSFRow)it.next();
            int rownum = row.getRowNum();
            
            if(rownum < startRow) {
                // ZSS-302: before skipping, must remove rows of destination that don't be replaced   
                if (canRemoveRow(startRow, endRow, n, rownum)) {
                    rowsToRemove.add(row);
                } 
            	continue;
            }
            
            final int newrownum = rownum + n;
            final boolean inbound = 0 <= newrownum && newrownum <= maxrow;
            if (!inbound) {
            	rowsToRemove.add(row); 
            	continue;
            }
            
            if (!copyRowHeight) {
                row.setHeight((short)-1);
            }

            if (canRemoveRow(startRow, endRow, n, rownum)) {
            	rowsToRemove.add(row);
            } else if (rownum >= startRow && rownum <= endRow) {
                new XSSFRowHelper(row).shift(n);
            }
        }
        
		// ZSS-419: remove a row not only remove it from map but also from real sheet XML
		for(Row row : rowsToRemove) {
			removeRow(row);
		}
        
        //rebuild the _rows map ASAP or getRow(rownum) will be incorrect
        TreeMap<Integer, XSSFRow> rows = getRows();
        TreeMap<Integer, XSSFRow> map = new TreeMap<Integer, XSSFRow>();
        for(XSSFRow r : rows.values()) {
            map.put(r.getRowNum(), r);
        }
        setRows(map);
        
        //handle inserted rows
        if (srcRow != null) {
        	final int row2 = Math.min(startRow + n - 1, SpreadsheetVersion.EXCEL2007.getLastRowIndex());
        	for ( int rownum = startRow; rownum <= row2; ++rownum) {
        		XSSFRow row = getRow(rownum);
        		if (row == null) {
        			row = createRow(rownum); 
        		}
        		row.setHeight(srcHeight); //height
//        		if (srcStyle != null) {
//        			row.setRowStyle((HSSFCellStyle)copyFromStyleExceptBorder(srcStyle));//style
//        		}
        		if (srcCells != null) {
	        		for (Entry<Integer, Cell> cellEntry : srcCells.entrySet()) {
	        			final Cell srcCell = cellEntry.getValue();
        				final CellStyle cellStyle = srcCell.getCellStyle();
        				final int c = cellEntry.getKey().intValue();
	        			Cell cell = row.getCell(c);
	        			if (cell == null) {
	        				cell = row.createCell(c);
	        			}
	        			cell.setCellStyle(BookHelper.copyFromStyleExceptBorder(getBook(), cellStyle));
	        		}
        		}
        	}
        }
        
        // Shift Hyperlinks which have been moved
        shiftHyperlinks(startRow, maxrow, n, 0, maxcol, 0);
        
        //special case1: endRow < startRow
        //special case2: (endRow - startRow + 1) < ABS(n)
        if (n < 0) {
        	if (endRow < startRow) { //special case1
	    		final int orgStartRow = startRow + n;
	            for ( int rowNum = orgStartRow; rowNum >= orgStartRow && rowNum <= endRow && rowNum >= 0 && rowNum <= maxrow; ++rowNum) {
	                final XSSFRow row = getRow( rowNum );
	                if (row != null) {
	                	removeRow(row);
	                }
	            }
	            removeHyperlinks(orgStartRow, endRow, 0, maxcol);
            } else if (clearRest) { //special case 2
            	final int orgStartRow = endRow + n + 1;
            	if (orgStartRow <= startRow) {
    	            for ( int rowNum = orgStartRow; rowNum >= orgStartRow && rowNum <= startRow && rowNum >= 0 && rowNum <= maxrow; ++rowNum) {
    	                final XSSFRow row = getRow( rowNum );
    	                if (row != null) {
    	                	removeRow(row);
    	                }
    	            }
    	            removeHyperlinks(orgStartRow, startRow, 0, maxcol);
            	}
        	}
        }

        //update named ranges
        if (startRow <= endRow) {
        	final XSSFWorkbook wb = getWorkbook();
        	int sheetIndex = wb.getSheetIndex(this);
        	final PtgShifter shifter = new PtgShifter(sheetIndex, startRow, endRow, n, 0, maxcol, 0, SpreadsheetVersion.EXCEL2007);
        	updateNamedRanges(wb, shifter);
        }
        
        return shiftedRanges;
    }
    
    private void updateNamedRanges(XSSFWorkbook wb, PtgShifter shifter) {
        XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.create(wb);
        for (int i = 0; i < wb.getNumberOfNames(); i++) {
            XSSFName name = wb.getNameAt(i);
            String formula = name.getRefersToFormula();
            int sheetIndex = name.getSheetIndex();

            // 20120904 samchuang@zkoss.org: ZSS-153, user define formula name range doesn't need to adjust range
            if (formula != null) {
                Ptg[] ptgs = FormulaParser.parse(formula, fpb, FormulaType.NAMEDRANGE, sheetIndex);
                if (shifter.adjustFormula(ptgs, sheetIndex)) {
                    String shiftedFmla = FormulaRenderer.toFormulaString(fpb, ptgs);
                    name.setRefersToFormula(shiftedFmla);
                }	
            }
        }
    }

    private boolean canRemoveRow(int startRow, int endRow, int n, int rownum) {
        if (rownum >= (startRow + n) && rownum <= (endRow + n)) {
            if (n > 0 && rownum > endRow) {
                return true;
            }
            else if (n < 0 && rownum < startRow) {
                return true;
            }
        }
        return false;
    }

    /**
     * Shift Hyperlink of the specified range.
     */
    //20130808 dennis remove hyperlink
    public void removeHyperlink(int row,int column){
    	removeHyperlinks(row,row,column,column);
    }
	
    //20130930 dennis remove hyperlinks
    private void shiftHyperlinks(int tRow, int bRow, int nRow, int lCol, int rCol, int nCol) {   	
    	Set<XSSFHyperlink> move = new HashSet<XSSFHyperlink>();
    	Set<XSSFHyperlink> remove = new HashSet<XSSFHyperlink>();
    	
    	//target area
    	int tRow2 = tRow+nRow;
    	int bRow2 = bRow+nRow;
    	int lCol2 = lCol+nCol;
    	int rCol2 = rCol+nCol;
    	
    	for(XSSFHyperlink hyperlink : hyperlinks) {
    		int r = hyperlink.getFirstRow();
    		int c = hyperlink.getFirstColumn();
            if(r>=tRow && r<=bRow && c>=lCol && c<=rCol) {
            	move.add(hyperlink);
            }
            if(r>=tRow2 && r<=bRow2 && c>=lCol2 && c<=rCol2) {
            	remove.add(hyperlink);
            }
        }
    	remove.removeAll(move);
    	
    	//remove the link in the destination area
    	hyperlinks.removeAll(remove);
    	
    	//move the link from src to destination area
    	for(XSSFHyperlink hyperlink : move) {
    		int row = hyperlink.getFirstRow()+nRow;
    		int col = hyperlink.getFirstColumn()+nCol;
    		hyperlink.setFirstRow(row);
    		hyperlink.setFirstColumn(col);
    	}
    }
    //20130930 dennis remove hyperlinks
    private void removeHyperlinks(int tRow, int bRow, int lCol, int rCol) {
    	HashSet<XSSFHyperlink> remove = new HashSet<XSSFHyperlink>();
    	for(XSSFHyperlink hyperlink : hyperlinks) {
    		int r = hyperlink.getFirstRow();
    		int c = hyperlink.getFirstColumn();
            if(r>=tRow && r<=bRow && c>=lCol && c<=rCol) {
                remove.add(hyperlink);
            }
        }
    	
    	if(remove.size()>0){
    		try{
    			//has also remove the relation 
            	PackageRelationshipCollection hyperRels =
                         getPackagePart().getRelationshipsByType(XSSFRelation.SHEET_HYPERLINKS.getRelation());
            	for(XSSFHyperlink hyperlink : remove) {
            		CTHyperlink cth = hyperlink.getCTHyperlink(); 
            		if(cth.isSetId()){//it is possible no ref id(e.g. DOC type)
            			String rid = cth.getId();
            			getPackagePart().removeRelationship(rid);
                    }
            	}
        	} catch (InvalidFormatException e){
                throw new POIXMLException(e);
            }    		
    		hyperlinks.removeAll(remove);
    	}
    }
    
    //20100520, henrichen@zkoss.org: Shift rows of a range
    /**
     * Shifts rows of a range between startRow and endRow n number of rows in the boundary of left column(lCol) and right column(rCol).
     * If you use a negative number, it will shift rows up.
     * Code ensures that rows don't wrap around
     *
     * <p>
     * Additionally shifts merged regions that are completely defined in these
     * rows (ie. merged 2 cells on a row to be shifted).
     * <p>
     * TODO Might want to add bounds checking here
     * @param startRow the row to start shifting
     * @param endRow the row to end shifting
     * @param n the number of rows to shift
     * @param lCol left boundary column
     * @param rCol right boundary column
     * @param copyRowHeight whether to copy the row height during the shift
     * @param resetOriginalRowHeight whether to set the original row's height to the default
     * @param moveComments whether to move comments at the same time as the cells they are attached to
     * @param clearRest whether clear the rest row after shifted endRow (meaningful only when n < 0)
     * @param copyOrigin copy format from the above/below row for the inserted rows(meaningful only when n > 0)
     * @return List of shifted merge ranges 
     */
    public List<CellRangeAddress[]> shiftRowsRange(int startRow, int endRow, int n, int lCol, int rCol,
            boolean copyRowHeight, boolean resetOriginalRowHeight, boolean moveComments, boolean clearRest, int copyOrigin) {

        final int maxrow = SpreadsheetVersion.EXCEL2007.getLastRowIndex();
        final int maxcol = SpreadsheetVersion.EXCEL2007.getLastColumnIndex();
        
	    // ZSS-418: moving a comment must move both comment and shape positions
	    // Otherwise, Excel will fail to load such file
        if(moveComments) {
            shiftComments(startRow, maxrow, lCol, rCol, n, 0); // take care: insert, delete and drag move
        }

    	//prepare source format row
    	final int srcRownum = n <= 0 ? -1 : copyOrigin == XRange.FORMAT_RIGHTBELOW ? startRow : copyOrigin == XRange.FORMAT_LEFTABOVE ? startRow - 1 : -1;
    	final XSSFRow srcRow = srcRownum >= 0 ? getRow(srcRownum) : null;
    	final Map<Integer, Cell> srcCells = srcRow != null ? BookHelper.copyRowCells(srcRow, lCol, rCol) : null;
    	final short srcHeight = srcRow != null ? srcRow.getHeight() : -1;
//    	final XSSFCellStyle srcStyle = srcRow != null ? srcRow.getRowStyle() : null;
    	
        final List<CellRangeAddress[]> shiftedRanges = BookHelper.shiftMergedRegion(this, startRow, lCol, maxrow, rCol, n, false);
        
        if (endRow < 0) {
        	endRow = maxrow;
        }
        final boolean wholeRow = lCol == 0 && rCol == maxcol; 
    	
        List<Row> rowsToRemove = new ArrayList<Row>(); // ZSS-419: queue row for removing later, remove can't be perform in this loop 
        final List<int[]> removePairs = new ArrayList<int[]>(); //row spans to be removed 
        final TreeMap<Integer, TreeMap<Integer, XSSFCell>> rowCells = new TreeMap<Integer, TreeMap<Integer, XSSFCell>>();
        int expectRownum = startRow; //handle sparse rows which might override destination row
        for (Iterator<Row> it = rowIterator() ; it.hasNext() ; ) { //TODO use submap between startRow and endRow
            XSSFRow row = (XSSFRow)it.next();
            int rownum = row.getRowNum();
            if (rownum < startRow) {
				// ZSS-389: before skipping. must remove cells of current row if current row will be deleted.
				if(canRemoveRow(startRow, endRow, n, rownum)) {
					row.getCells().subMap(lCol, rCol + 1).clear();
				}
            	continue;
            }
            if (rownum > endRow) break; //no more
            
            final int newRownum = rownum + n;
            if (rownum > expectRownum) { //sparse row between expectRownum(inclusive) and current row(exclusive), to be removed
            	addRemovePair(removePairs, expectRownum + n, newRownum);
            }
            expectRownum = rownum + 1;
            
            final boolean inbound = 0 <= newRownum && newRownum <= maxrow;
            if (!inbound) {
            	rowsToRemove.add(row);
            	continue;
            }
            
            if (wholeRow) {
	            if (!copyRowHeight) {
	                row.setHeight((short)-1);
	            }
	            if (canRemoveRow(startRow, endRow, n, rownum)) {
	                rowsToRemove.add(row);
	            }
	            else if (rownum >= startRow && rownum <= endRow) {
	                new XSSFRowHelper(row).shift(n);
	            }
            } else {
            	SortedMap<Integer, XSSFCell> oldCells = row.getCells().subMap(Integer.valueOf(lCol), Integer.valueOf(rCol+1));
            	if (!oldCells.isEmpty()) {
            		TreeMap<Integer, XSSFCell> cells = new TreeMap<Integer, XSSFCell>(oldCells);
            		rowCells.put(newRownum, cells);
            		for (Cell cell : cells.values()) {
            			row.removeCell(cell);
            		}
            	}
            }
        }
        
		// ZSS-419: remove a row not only remove it from map but also from real sheet XML
		for(Row row : rowsToRemove) {
			removeRow(row);
		}

        //rebuild rows ASAP or the getRow(rownum) will be incorrect
        if (wholeRow) {
	        TreeMap<Integer, XSSFRow> map = new TreeMap<Integer, XSSFRow>();
	        TreeMap<Integer, XSSFRow> rows = getRows();
	        for(XSSFRow r : rows.values()) {
	            map.put(r.getRowNum(), r);
	        }
	        setRows(map);
        }
        
        //sparse row between expectRownum(inclusive) to endRow+1(exclusive), to be removed
    	addRemovePair(removePairs, expectRownum + n, endRow + 1 + n);
    	
        //really remove rows
        if (wholeRow) {
        	for(int[] pair : removePairs) {
        		final int start = Math.max(0, pair[0]);
        		final int end = Math.min(SpreadsheetVersion.EXCEL2007.getLastRowIndex() + 1, pair[1]);
        		for(int j=start; j < end; ++j) {
        			Row row = getRow(j);
        			if (row != null) {
        				removeRow(row);
        			}
        		}
        	}
        } else { //clear cells between lCol and rCol
        	for(int[] pair : removePairs) {
        		final int start = Math.max(0, pair[0]);
        		final int end = pair[1];
        		for(int j=start; j < end; ++j) {
        			Row row = getRow(j);
        			if (row != null) {
        				removeCells(row, lCol, rCol);
        			}
        		}
        	}
        }
        
        //really update the row's cells
        for (Entry<Integer, TreeMap<Integer, XSSFCell>> entry : rowCells.entrySet()) {
        	final int rownum = entry.getKey().intValue();
        	final TreeMap<Integer, XSSFCell> cells = entry.getValue();
        	XSSFRow row = getRow(rownum);
        	if (row == null) {
        		row = createRow(rownum);
        	} else {
        		removeCells(row, lCol, rCol);
        	}
        	for(Entry<Integer, XSSFCell> cellentry : cells.entrySet()) {
        		final int colnum = cellentry.getKey().intValue();
        		final XSSFCell srcCell = cellentry.getValue();
        		BookHelper.assignCell(srcCell, row.createCell(colnum));
        	}
        }
        
        //handle inserted rows
        if (srcRow != null) {
        	final int row2 = Math.min(startRow + n - 1, SpreadsheetVersion.EXCEL2007.getLastRowIndex());
        	for ( int rownum = startRow; rownum <= row2; ++rownum) {
        		XSSFRow row = getRow(rownum);
        		if (row == null) {
        			row = createRow(rownum); 
        		}
        		row.setHeight(srcHeight); //height
//        		if (srcStyle != null) {
//        			row.setRowStyle((HSSFCellStyle)copyFromStyleExceptBorder(srcStyle));//style
//        		}
        		if (srcCells != null) {
	        		for (Entry<Integer, Cell> cellEntry : srcCells.entrySet()) {
	        			final Cell srcCell = cellEntry.getValue();
        				final CellStyle cellStyle = srcCell.getCellStyle();
        				final int c = cellEntry.getKey().intValue();
	        			Cell cell = row.getCell(c);
	        			if (cell == null) {
	        				cell = row.createCell(c);
	        			}
	        			cell.setCellStyle(BookHelper.copyFromStyleExceptBorder(getBook(), cellStyle));
	        		}
        		}
        	}
        }
        
        // Shift Hyperlinks which have been moved
        shiftHyperlinks(startRow, maxrow, n, lCol, rCol, 0);
        
        //special case1: endRow < startRow
        //special case2: (endRow - startRow + 1) < ABS(n)
        if (n < 0) {
        	if (endRow < startRow) { //special case1
	    		final int orgStartRow = startRow + n;
	            for ( int rowNum = orgStartRow; rowNum >= orgStartRow && rowNum <= endRow && rowNum >= 0 && rowNum <= maxrow; ++rowNum) {
	                final XSSFRow row = getRow( rowNum );
	                if (row != null) {
	                	removeRow(row);
	                }
	            }
	            removeHyperlinks(orgStartRow, endRow, 0, maxcol);
            } else if (clearRest) { //special case 2
            	final int orgStartRow = endRow + n + 1;
            	if (orgStartRow <= startRow) {
    	            for ( int rowNum = orgStartRow; rowNum >= orgStartRow && rowNum <= startRow && rowNum >= 0 && rowNum <= maxrow; ++rowNum) {
    	                final XSSFRow row = getRow( rowNum );
    	                if (row != null) {
    	                	removeRow(row);
    	                }
    	            }
    	            removeHyperlinks(orgStartRow, startRow, 0, maxcol);
            	}
        	}
        }
        
        //update named ranges
        if (startRow <= endRow) {
        	final XSSFWorkbook wb = getWorkbook();
	        int sheetIndex = wb.getSheetIndex(this);
	        final PtgShifter shifter = new PtgShifter(sheetIndex, startRow, endRow, n, 0, maxcol, 0, SpreadsheetVersion.EXCEL2007);
	        updateNamedRanges(wb, shifter);
        }

        return shiftedRanges;
	}
    private void addRemovePair(List<int[]> removePairs, int start, int end) {
    	if (start < 0) {
    		start = 0;
    	}
    	if (end > (getLastRowNum() + 1)) {
    		end = getLastRowNum() + 1;
    	}
    	if (start < end) {
    		removePairs.add(new int[] {start, end});
    	}
    }
    private void removeCells(Row row, int lCol, int rCol) {
    	for(int j = lCol; j <= rCol; ++j) {
    		Cell cell = row.getCell(j);
    		if (cell != null) {
    			row.removeCell(cell);
    		}
    	}
    }
    /**
     * Shifts columns between startCol and endCol n number of columns.
     * If you use a negative number, it will shift columns left.
     * Code ensures that columns don't wrap around
     *
     * <p>
     * Additionally shifts merged regions that are completely defined in these
     * columns (ie. merged 2 cells on a column to be shifted).
     * <p>
     * @param startCol the column to start shifting
     * @param endCol the column to end shifting; -1 means using the last available column number 
     * @param n the number of rows to shift
     * @param copyColWidth whether to copy the column width during the shift
     * @param resetOriginalColWidth whether to set the original column's height to the default
     * @param moveComments whether to move comments at the same time as the cells they are attached to
     * @param clearRest whether clear cells after the shifted endCol (meaningful only when n < 0)
     * @param copyOrigin copy format from the left/right column for the inserted column(meaningful only when n > 0)
     * @return List of shifted merge ranges
     */
    public List<CellRangeAddress[]> shiftColumnsOnly(int startCol, int endCol, int n,
            boolean copyColWidth, boolean resetOriginalColWidth, boolean moveComments, boolean clearRest, int copyOrigin) {

        final int maxrow = SpreadsheetVersion.EXCEL2007.getLastRowIndex();
        final int maxcol = SpreadsheetVersion.EXCEL2007.getLastColumnIndex();

        // Move comments from the source column to the
        // destination column. Note that comments can
		// exist for cells which are null
		if(moveComments) {
		    // ZSS-418: moving a comment must move both comment and shape positions
		    // Otherwise, Excel will fail to load such file
            shiftComments(0, maxrow, startCol, maxcol, 0, n);
		}
    	
    	//prepared inserting column format
    	final int srcCol = n <= 0 ? -1 : copyOrigin == XRange.FORMAT_RIGHTBELOW ? startCol : copyOrigin == XRange.FORMAT_LEFTABOVE ? startCol - 1 : -1; 
    	final CellStyle colStyle = srcCol >= 0 ? getColumnStyle(srcCol) : null;
    	final int colWidth = srcCol >= 0 ? getColumnWidth(srcCol) : -1; 
    	final Map<Integer, Cell> cells = srcCol >= 0 ? new HashMap<Integer, Cell>() : null;
    	
    	int maxColNum = -1;
        for (Iterator<Row> it = rowIterator() ; it.hasNext() ; ) {
            XSSFRow row = (XSSFRow)it.next();
            int rowNum = row.getRowNum();
            
            if (endCol < 0) {
	            final int colNum = row.getLastCellNum() - 1;
	            if (colNum > maxColNum)
	            	maxColNum = colNum;
            }
            
            if (cells != null) {
           		final Cell cell = row.getCell(srcCol);
           		if (cell != null) {
           			cells.put(Integer.valueOf(rowNum), cell);
           		}
            }
            
            shiftCells(row, startCol, endCol, n, clearRest);
        }
        
        final List<CellRangeAddress[]> shiftedRanges = BookHelper.shiftMergedRegion(this, 0, startCol, maxrow, maxcol, n, true);
        
        if (endCol < 0) {
        	endCol = maxColNum;
        }
        if (n > 0) {
	        if (startCol > endCol) { //nothing to do
	        	return shiftedRanges;
	        }
        } else {
        	if ((startCol + n) > endCol) { //nothing to do
	        	return shiftedRanges;
        	}
        }
        
        //TODO handle the page breaks
        //?
        
        // Fix up column width if required
        int s, inc;
        if (n < 0) {
            s = startCol;
            inc = 1;
        } else {
            s = endCol;
            inc = -1;
        } 

        if (copyColWidth || resetOriginalColWidth) {
        	final int defaultColumnWidth = getDefaultColumnWidth();
	        for ( int colNum = s; colNum >= startCol && colNum <= endCol && colNum >= 0 && colNum <= maxcol; colNum += inc ) {
	        	final int newColNum = colNum + n;
		        if (copyColWidth) {
		            setColumnWidth(newColNum, getColumnWidth(colNum));
		        }
		        if (resetOriginalColWidth) {
		            setColumnWidth(colNum, defaultColumnWidth);
		        }
	        }
        }

        //handle inserted columns
        if (srcCol >= 0) {
        	final int col2 = Math.min(startCol + n - 1, maxcol);
        	for (int col = startCol; col <= col2 ; ++col) {
        		//copy the column width
        		setColumnWidth(col, colWidth);
        		if (colStyle != null) {
        			setDefaultColumnStyle(col, BookHelper.copyFromStyleExceptBorder(getBook(), colStyle));
        		}
        	}
        	if (cells != null) {
		        for (Entry<Integer, Cell> cellEntry : cells.entrySet()) {
		            final XSSFRow row = getRow(cellEntry.getKey().intValue());
		            final Cell srcCell = cellEntry.getValue();
		            final CellStyle srcStyle = srcCell.getCellStyle();
		        	for (int col = startCol; col <= col2; ++col) {
		        		Cell dstCell = row.getCell(col);
		        		if (dstCell == null) {
		        			dstCell = row.createCell(col);
		        		}
		        		dstCell.setCellStyle(BookHelper.copyFromStyleExceptBorder(getBook(), srcStyle));
		        	}
		        }
        	}
        }
        
        // Shift Hyperlinks which have been moved
        shiftHyperlinks(0, maxrow, 0, startCol, maxcol, n);
        
        //special case1: endCol < startCol
        //special case2: (endCol - startCol + 1) < ABS(n)
        if (n < 0) {
        	if (endCol < startCol) { //special case1
	    		final int replacedStartCol = startCol + n;
	    		removeHyperlinks(0, maxrow, replacedStartCol, endCol);
            } else if (clearRest) { //special case 2
            	final int replacedStartCol = endCol + n + 1;
            	if (replacedStartCol <= startCol) {
    	    		removeHyperlinks(0, maxrow, replacedStartCol, startCol);
            	}
        	}
        }
        
        // Update any formulas on this sheet that point to
        // columns which have been moved
        if (startCol <= endCol) {
        	XSSFWorkbook book = getWorkbook();
        	int sheetIndex = book.getSheetIndex(this);
        	PtgShifter shifter = new PtgShifter(sheetIndex, 0, maxrow, 0, startCol, endCol, n, SpreadsheetVersion.EXCEL2007);
        	updateNamedRanges(book, shifter);
        }
        
        return shiftedRanges;
    }
    
    //20100916, henrichen@zkoss.org
    /**
     * Shifts cells between startColumn and endColumn n number of columns.
     * If you use a negative number, it will shift columns left.
     * Code ensures that columns don't wrap around
     *
     * @param startCol the column to start shifting
     * @param endCol the column to end shifting
     * @param n the number of columns to shift
     * @param clearRest whether clear the rest cells after the shifted endCol
     */
    private void shiftCells(XSSFRow row, int startCol, int endCol, int n, boolean clearRest) {
        final int maxcol = SpreadsheetVersion.EXCEL2007.getLastColumnIndex();
    	if (endCol < 0) {
    		endCol = row.getLastCellNum() - 1;
    	}
    	if (n > 0) {
    		if (endCol < startCol) { //nothing to do
    			return;
    		}
    	} else {
    		if (endCol < (startCol + n)) { //nothing to do
    			return;
    		}
    	}
    	
    	List<XSSFCell> cellsToRemove = new ArrayList<XSSFCell>(); // ZSS-179: queue cells for removing later
        final List<int[]> removePairs = new ArrayList<int[]>(); //column spans to be removed 
        final Set<XSSFCell> rowCells = new HashSet<XSSFCell>();
        int expectColnum = startCol; //handle sparse columns which might override destination column
        if(endCol >= startCol) { // ZSS-245: some row's last cell might be inside the columns that will be removed
        	for (Iterator<XSSFCell> it = row.getCells().subMap(startCol, endCol+1).values().iterator(); it.hasNext(); ) {
        		XSSFCell cell = it.next();
        		int colnum = cell.getColumnIndex();
        		
        		final int newColnum = colnum + n;
        		if (colnum > expectColnum) { //sparse column between expectColnum(inclusive) and current column(exclusive), to be removed
        			addRemovePair(removePairs, row, expectColnum + n, newColnum);
        		}
        		expectColnum = colnum + 1;
        		
        		cellsToRemove.add(cell);
        		final boolean inbound = 0 <= newColnum && newColnum <= maxcol;
        		if (!inbound) {
        			notifyCellShifting(cell);
        			continue;
        		}
        		rowCells.add(cell);
        	}
        }
        for(XSSFCell c : cellsToRemove) {
        	row.removeCell(c);
        }
    	
    	addRemovePair(removePairs, row, expectColnum + n, endCol + 1 + n);
    	
    	//remove those not existing cells
    	for(int[] pair : removePairs) {
    		final int start = Math.max(0, pair[0]);
    		final int end = Math.min(maxcol + 1, pair[1]);
    		for(int j=start; j < end; ++j) {
    			Cell cell = row.getCell(j);
    			if (cell != null) {
    				row.removeCell(cell);
    			}
    		}
    	}
    	
    	//update the cells
    	for(XSSFCell srcCell: rowCells) {
    		BookHelper.assignCell(srcCell, row.createCell(srcCell.getColumnIndex()+n));
    	}
    	
        //special case1: endCol < startCol
        //special case2: (endCol - startCol + 1) < ABS(n)
        if (n < 0) {
        	if (endCol < startCol) { //special case1
	    		final int replacedStartCol = startCol + n;
	            for ( int colNum = replacedStartCol; colNum >= replacedStartCol && colNum <= endCol && colNum >= 0 && colNum <= maxcol; ++colNum) {
	            	final XSSFCell cell = row.getCell(colNum, Row.RETURN_NULL_AND_BLANK);
	            	if (cell != null) {
	            		row.removeCell(cell);
	            	}
	            }
            } else if (clearRest) { //special case 2
            	final int replacedStartCol = endCol + n + 1;
            	if (replacedStartCol <= startCol) {
    	            for ( int colNum = replacedStartCol; colNum >= replacedStartCol && colNum <= startCol && colNum >= 0 && colNum <= maxcol; ++colNum) {
    	            	final XSSFCell cell = row.getCell(colNum, Row.RETURN_NULL_AND_BLANK);
    	            	if (cell != null) {
    	            		row.removeCell(cell);
    	            	}
    	            }
            	}
        	}
        }
    }
    
    private void addRemovePair(List<int[]> removePairs, XSSFRow row, int start, int end) {
    	if (start < 0) {
    		start = 0;
    	}
    	if (end > row.getLastCellNum()) {
    		end = row.getLastCellNum(); //index returned already plus 1
    	}
    	if (start < end) {
    		removePairs.add(new int[] {start, end});
    	}
    }

    //20100916, henrichen@zkoss.org
    private void notifyCellShifting(XSSFCell cell){
        String msg = "Cell[rownum="+cell.getRowIndex()+", columnnum="+cell.getColumnIndex()+"] included in a multi-cell array formula. " +
                "You cannot change part of an array.";
        if(cell.isPartOfArrayFormulaGroup()){
            new XSSFCellHelper(cell).notifyArrayFormulaChanging(msg);
        }
    }
    
    //20100701, henrichen@zkoss.org: Shift columns of a range
    /**
     * Shifts columns of a range between startCol and endCol n number of columns in the boundary of top row(tRow) and bottom row(bRow).
     * If you use a negative number, it will shift columns left.
     * Code ensures that columns don't wrap around
     *
     * <p>
     * Additionally shifts merged regions that are completely defined in these
     * columns (ie. merged 2 cells on a column to be shifted) within the specified boundary rows.
     * <p>
     * @param startCol the column to start shifting
     * @param endCol the column to end shifting; -1 means using the last available column number 
     * @param n the number of rows to shift
     * @param tRow top boundary row index
     * @param bRow bottom boundary row index
     * @param copyColWidth whether to copy the column width during the shift
     * @param resetOriginalColWidth whether to set the original column's height to the default
     * @param moveComments whether to move comments at the same time as the cells they are attached to
     * @param clearRest whether clear cells after the shifted endCol (meaningful only when n < 0)
     * @param copyOrigin copy format from the left/right column for the inserted column(meaningful only when n > 0)
     * @return List of shifted merge ranges
     */
    public List<CellRangeAddress[]> shiftColumnsRange(int startCol, int endCol, int n, int tRow, int bRow,
            boolean copyColWidth, boolean resetOriginalColWidth, boolean moveComments, boolean clearRest, int copyOrigin) {

        final int maxrow = SpreadsheetVersion.EXCEL2007.getLastRowIndex();
        final int maxcol = SpreadsheetVersion.EXCEL2007.getLastColumnIndex();

        // Move comments from the source column to the
        //  destination column. Note that comments can
        //  exist for cells which are null
        if (moveComments) {
		    // ZSS-418: moving a comment must move both comment and shape positions
		    // Otherwise, Excel will fail to load such file
            shiftComments(tRow, bRow, startCol, maxcol, 0, n); // take care: insert, delete and drag move
        }
        
    	//prepared inserting column format
    	final int srcCol = n <= 0 ? -1 : copyOrigin == XRange.FORMAT_RIGHTBELOW ? startCol : copyOrigin == XRange.FORMAT_LEFTABOVE ? startCol - 1 : -1; 
    	final CellStyle colStyle = srcCol >= 0 ? getColumnStyle(srcCol) : null;
    	final int colWidth = srcCol >= 0 ? getColumnWidth(srcCol) : -1; 
    	final Map<Integer, Cell> cells = srcCol >= 0 ? new HashMap<Integer, Cell>() : null;
    	
    	int startRow = Math.max(tRow, getFirstRowNum());
    	int endRow = Math.min(bRow, getLastRowNum());
    	int maxColNum = -1;
    	if (startRow <= endRow) {
	        for (Iterator<XSSFRow> it = getRows().subMap(startRow, endRow+1).values().iterator(); it.hasNext() ; ) {
	            XSSFRow row = it.next();
	            int rowNum = row.getRowNum();
	            
	            if (endCol < 0) {
		            final int colNum = row.getLastCellNum() - 1;
		            if (colNum > maxColNum)
		            	maxColNum = colNum;
	            }
	            
	            if (cells != null) {
	           		final Cell cell = row.getCell(srcCol);
	           		if (cell != null) {
	           			cells.put(Integer.valueOf(rowNum), cell);
	           		}
	            }
	            
	            shiftCells(row, startCol, endCol, n, clearRest);
	        }
    	}
        
        final List<CellRangeAddress[]> shiftedRanges = BookHelper.shiftMergedRegion(this, tRow, startCol, bRow, maxcol, n, true);
    	
        if (endCol < 0) {
        	endCol = maxColNum;
        }
        if (n > 0) {
	        if (startCol > endCol) { //nothing to do
	        	return shiftedRanges;
	        }
        } else {
        	if ((startCol + n) > endCol) { //nothing to do
	        	return shiftedRanges;
        	}
        }
        
        final boolean wholeColumn = tRow == 0 && bRow == maxrow; 
        if (wholeColumn) {
	        //TODO handle the page breaks
	        //?
        }
        
        // Fix up column width if required
        int s, inc;
        if (n < 0) {
            s = startCol;
            inc = 1;
        } else {
            s = endCol;
            inc = -1;
        } 

        if (wholeColumn && (copyColWidth || resetOriginalColWidth)) {
        	final int defaultColumnWidth = getDefaultColumnWidth();
	        for ( int colNum = s; colNum >= startCol && colNum <= endCol && colNum >= 0 && colNum <= maxcol; colNum += inc ) {
	        	final int newColNum = colNum + n;
		        if (copyColWidth) {
		            setColumnWidth(newColNum, getColumnWidth(colNum));
		        }
		        if (resetOriginalColWidth) {
		            setColumnWidth(colNum, defaultColumnWidth);
		        }
	        }
        }

        //handle inserted columns
        if (srcCol >= 0) {
        	final int col2 = Math.min(startCol + n - 1, maxcol);
        	if (wholeColumn) {
	        	for (int col = startCol; col <= col2 ; ++col) {
	        		//copy the column width
	        		setColumnWidth(col, colWidth);
	        		if (colStyle != null) {
	        			setDefaultColumnStyle(col, BookHelper.copyFromStyleExceptBorder(getBook(), colStyle));
	        		}
	        	}
        	}
        	if (cells != null) {
		        for (Entry<Integer, Cell> cellEntry : cells.entrySet()) {
		            final XSSFRow row = getRow(cellEntry.getKey().intValue());
		            final Cell srcCell = cellEntry.getValue();
		            final CellStyle srcStyle = srcCell.getCellStyle();
		        	for (int col = startCol; col <= col2; ++col) {
		        		Cell dstCell = row.getCell(col);
		        		if (dstCell == null) {
		        			dstCell = row.createCell(col);
		        		}
		        		dstCell.setCellStyle(BookHelper.copyFromStyleExceptBorder(getBook(), srcStyle));
		        	}
		        }
        	}
        }
        
        // Shift Hyperlinks which have been moved
        shiftHyperlinks(tRow, bRow, 0, startCol, maxcol, n);
        
        //special case1: endCol < startCol
        //special case2: (endCol - startCol + 1) < ABS(n)
        if (n < 0) {
        	if (endCol < startCol) { //special case1
	    		final int replacedStartCol = startCol + n;
	    		removeHyperlinks(0, maxrow, replacedStartCol, endCol);
            } else if (clearRest) { //special case 2
            	final int replacedStartCol = endCol + n + 1;
            	if (replacedStartCol <= startCol) {
    	    		removeHyperlinks(0, maxrow, replacedStartCol, startCol);
            	}
        	}
        }
        
        // Update any formulas on this sheet that point to
        // columns which have been moved
        if (startCol <= endCol) {
        	XSSFWorkbook book = getWorkbook();
	        int sheetIndex = book.getSheetIndex(this);
	        PtgShifter shifter = new PtgShifter(sheetIndex, 0, maxrow, 0, startCol, endCol, n, SpreadsheetVersion.EXCEL2007);
	        updateNamedRanges(book, shifter);
        }
        
        return shiftedRanges;
    }
    
    public List<CellRangeAddress[]> shiftBothRange(int tRow, int bRow, int nRow, int lCol, int rCol, int nCol, boolean moveComments) {

        // Move comments from the source column to the
        //  destination column. Note that comments can
        //  exist for cells which are null
		if(moveComments) {
		    // ZSS-418: moving a comment must move both comment and shape positions
		    // Otherwise, Excel will fail to load such file
			shiftComments(tRow, bRow, lCol, rCol, nRow, nCol);
		}
    	
        final int maxrow = SpreadsheetVersion.EXCEL2007.getLastRowIndex();
        final int maxcol = SpreadsheetVersion.EXCEL2007.getLastColumnIndex();
        final List<CellRangeAddress[]> shiftedRanges = BookHelper.shiftBothMergedRegion(this, tRow, lCol, bRow, rCol, nRow, nCol);
		
    	int startRow = Math.max(tRow, getFirstRowNum());
    	int endRow = Math.min(bRow, getLastRowNum());
        if (nRow > 0) {
	        if (tRow > endRow ) { //nothing to do
	        	return shiftedRanges;
	        }
        } else {
        	if ((tRow + nRow) > endRow) { //nothing to do
	        	return shiftedRanges;
        	}
        }
        
        
        final List<int[]> removePairs = new ArrayList<int[]>(); //row spans to be removed 
        final TreeMap<Integer, TreeMap<Integer, XSSFCell>> rowCells = new TreeMap<Integer, TreeMap<Integer, XSSFCell>>();
        int expectRownum = tRow; //handle sparse rows which might override destination row
    	int maxColNum = -1;
    	if (startRow <= endRow) {
	        for (Iterator<XSSFRow> it = getRows().subMap(startRow, endRow+1).values().iterator(); it.hasNext() ; ) {
	            XSSFRow row = it.next();
	            int rownum = row.getRowNum();

	            final int newRownum = rownum + nRow;
	            if (newRownum > maxrow) { //nothing to do
	            	break;
	            }
	            if (rownum > expectRownum) { //sparse row between expectRownum(inclusive) and current row(exclusive), to be removed
	            	addRemovePair(removePairs, expectRownum + nRow, newRownum);
	            }
	            expectRownum = rownum + 1;
	            
            	SortedMap<Integer, XSSFCell> oldCells = row.getCells().subMap(Integer.valueOf(lCol), Integer.valueOf(rCol+1));
            	if (!oldCells.isEmpty()) {
            		TreeMap<Integer, XSSFCell> cells = new TreeMap<Integer, XSSFCell>(oldCells);
            		rowCells.put(newRownum, cells);
            		for(Cell cell : cells.values()) { //remove reference from row to the cell
            			row.removeCell(cell);
            		}
            	}
	        }
    	}
    	
    	//spare row between expectedRownum(inclusive) to endRow+1(exclusive), to be remove
    	addRemovePair(removePairs, expectRownum + nRow, endRow + 1 + nRow);
    	
    	//really remove rows of the target
    	final int tgtlCol = Math.max(0, lCol + nCol);
    	final int tgtrCol = Math.min(maxcol, rCol + nCol);
    	for(int[] pair : removePairs) {
    		final int start = Math.max(0, pair[0]);
    		final int end = pair[1];
    		for(int j=start; j < end; ++j) {
    			Row row = getRow(j);
    			if (row != null) {
    				removeCells(row, tgtlCol, tgtrCol);
    			}
    		}
    	}

        //really update the row's cells
        for (Entry<Integer, TreeMap<Integer, XSSFCell>> entry : rowCells.entrySet()) {
        	final int rownum = entry.getKey().intValue();
        	final TreeMap<Integer, XSSFCell> cells = entry.getValue();
        	XSSFRow row = getRow(rownum);
        	if (row == null) {
        		row = createRow(rownum);
        	} else {
        		removeCells(row, tgtlCol, tgtrCol);
        	}
        	for(Entry<Integer, XSSFCell> cellentry : cells.entrySet()) {
        		final int colnum = cellentry.getKey().intValue() + nCol;
        		if (colnum < 0) { //out of bound
        			continue;
        		}
        		if (colnum > maxcol) {
        			break;
        		}
        		final XSSFCell srcCell = cellentry.getValue();
        		BookHelper.assignCell(srcCell, row.createCell(colnum));
        	}
        }
    	
        // Shift Hyperlinks which have been moved
        shiftHyperlinks(tRow, bRow, nRow, lCol, rCol, nCol);
        
        // Update any formulas on this sheet that point to
        // columns which have been moved
        if (tRow <= bRow && lCol <= rCol) {
        	XSSFWorkbook book = getWorkbook();
	        int sheetIndex = book.getSheetIndex(this);
	        PtgShifter shifter = new PtgShifter(sheetIndex, tRow, bRow, nRow, lCol, rCol, nCol, SpreadsheetVersion.EXCEL2007);
	        updateNamedRanges(book, shifter);
        }
        
        return shiftedRanges;
    }
    
    public boolean isFreezePanes() {
        final CTWorksheet ctsheet = getCTWorksheet();
        final CTSheetViews views = ctsheet != null ? ctsheet.getSheetViews() : null;
        final List<CTSheetView> viewList = views != null ? views.getSheetViewList() : null; 
        final CTSheetView view = viewList != null && !viewList.isEmpty() ? viewList.get(0) : null;
    	final CTPane pane = view != null ? view.getPane() : null;
    	if (pane == null) {
    		return false;
    	} else {
    		return pane.getState() == STPaneState.FROZEN; 
    	}
    }

    private XSSFDrawing _patriarch = null;
    public XSSFDrawing getDrawingPatriarch() {
    	if (_patriarch == null) {
	        for(POIXMLDocumentPart dr : getRelations()){
	            if(dr instanceof XSSFDrawing){
	            	_patriarch = (XSSFDrawing) dr;
	            	break;
	            }
	        }
    	}
    	return _patriarch;
    }
    
    //--SheetCtrl--//
    private volatile SheetCtrl _sheetCtrl;
    private SheetCtrl getSheetCtrl() {
    	SheetCtrl ctrl = _sheetCtrl;
    	if (ctrl == null) {
    		synchronized(this) {
    			ctrl = _sheetCtrl;
    			if (ctrl == null) {
    				String clsnm = Library.getProperty("org.zkoss.zss.model.impl.SheetCtrl.class");
    				if (clsnm == null) {
    					clsnm = "org.zkoss.zss.model.impl.SheetCtrlImpl";
    				}
    				try {
						ctrl = _sheetCtrl = (SheetCtrl) Classes.newInstanceByThread(clsnm, new Class[] {XBook.class, XSheet.class}, new Object[] {getBook(), this});
					} catch (Exception e) {
						ctrl = _sheetCtrl = new SheetCtrlImpl(getBook(), this); 
					}
    			}
    		}
    	}
    	return ctrl;
    }
	@Override
	public void evalAll() {
		getSheetCtrl().evalAll();
	}

	@Override
	public boolean isEvalAll() {
		return getSheetCtrl().isEvalAll();
	}
	@Override
	public String getUuid() {
		return getSheetCtrl().getUuid();
	}

	@Override
	public void addMerged(CellRangeAddress addr) {
		getSheetCtrl().addMerged(addr);
	}

	@Override
	public void deleteMerged(CellRangeAddress addr) {
		getSheetCtrl().deleteMerged(addr);
	}
	@Override
	public CellRangeAddress getMerged(int row, int col) {
		return getSheetCtrl().getMerged(row, col);
	}
	@Override
	public void initMerged() {
		getSheetCtrl().initMerged();
	}
	@Override
	public DrawingManager getDrawingManager() {
		return getSheetCtrl().getDrawingManager();
	}
	@Override
	public void whenRenameSheet(String oldname, String newname) {
		getSheetCtrl().whenRenameSheet(oldname, newname);
		//handle formula reference
		for(Row row : this) {
			if (row != null) {
				for (Cell cell : row) {
					if (cell != null) {
						((XSSFCell) cell).whenRenameSheet(oldname, newname);
					}
				}
			}
		}
	}
	
	// ZSS-397: remove drawing part
	@Override
	public void removeDrawingPatriarch(Drawing drawing) {
		removeRelation((XSSFDrawing)drawing, true);
		worksheet.unsetDrawing();
	}
	
	
	// ZSS-418: move comments completely 
	/**
	 * To move comments according to specific range.
	 * All comments in destination will be removed.
	 * @param nRow move n rows, negative number means inverse direction.
	 * @param nCol move n columns, negative number means inverse direction.
	 */
	private void shiftComments(int row, int lastRow, int col, int lastCol, int nRow, int nCol) {

		if(nRow == 0 && nCol == 0) { // just in case
			return;
		}
		CommentsTable sheetComments = getCommentsTable(false);
		if(sheetComments == null) {
			return;
		}

		// collect comments in source range. we need get XSSFComment from XML beans and move comment through it
		Map<String, XSSFComment> src = getComments(sheetComments, row, lastRow, col, lastCol);

		// collect comments in destination range
		Map<String, XSSFComment> dest = getComments(sheetComments, row + nRow, lastRow + nRow, col + nCol,
				lastCol + nCol);

		// filter comments are also in source range (source overlaps destination range)
		Iterator<Entry<String, XSSFComment>> iter = dest.entrySet().iterator();
		while(iter.hasNext()) {
			if(src.containsKey(iter.next().getKey())) {
				iter.remove();
			}
		}

		// remove comments in destination
		for(XSSFComment c : dest.values()) {
			removeCellComment(c.getRow(), c.getColumn());
		}

		// when moving comments, the position CAN'T be duplicated.
		// So, we must sort them first.
		List<XSSFComment> ordered = new ArrayList<XSSFComment>(src.values());
		Collections.sort(ordered, new Comparator<XSSFComment>() {
			@Override
			public int compare(XSSFComment a, XSSFComment b) {
				if(a.getRow() != b.getRow()) {
					return a.getRow() - b.getRow();
				} else {
					return a.getColumn() - b.getColumn();
				}
			}
		});
		
		// move source comments to destination (they are just points)
		if(nRow != 0) {
			// decide forward or backward
			int index = nRow < 0 ? 0 : ordered.size() -1;
			int step = nRow < 0 ? 1 : -1;
			for(int count = 0 ; count < ordered.size() ; ++count, index = index + step) {
				XSSFComment comment = ordered.get(index);
				comment.setRow(comment.getRow() + nRow);
			}
		}
		if(nCol != 0) {
			// decide forward or backward
			int index = nCol < 0 ? 0 : ordered.size() -1;
			int step = nCol < 0 ? 1 : -1;
			for(int count = 0 ; count < ordered.size() ; ++count, index = index + step) {
				XSSFComment comment = ordered.get(index);
				comment.setColumn(comment.getColumn() + nCol);
			}
		}

		// remove comments that are out of bounds
		int maxRow = SpreadsheetVersion.EXCEL2007.getLastRowIndex();
		int maxCol = SpreadsheetVersion.EXCEL2007.getLastColumnIndex();
		for(XSSFComment c : src.values()) {
			if(!isInside(0, maxRow, 0, maxCol, c.getRow(), c.getColumn())) {
				removeCellComment(c.getRow(), c.getColumn());
			}
		}
	}
	
	private Map<String, XSSFComment> getComments(CommentsTable commentsTable, int row, int lastRow, int col, int lastCol) {
		Map<String, XSSFComment> comments = new HashMap<String, XSSFComment>();
		for(CTComment ctc : commentsTable.getCTComments().getCommentList().getCommentList()) {
			CellReference ref = new CellReference(ctc.getRef());
			if(isInside(row, lastRow, col, lastCol, ref.getRow(), ref.getCol())) {
				comments.put(ctc.getRef(), getCellComment(ref.getRow(), ref.getCol()));
			}
		}
		return comments;
	}

	private boolean isInside(int row, int lastRow, int col, int lastCol, int r, int c) {
		return row <= r && r <= lastRow && col <= c && c <= lastCol;
	}
	
	//ZSS-256, add api to get correct configuraed column index
	public int getMaxConfiguredColumn() {
		int max = -1;
		CTWorksheet worksheet = getCTWorksheet();
		CTCols[] colsArrays = worksheet.getColsArray(); // might be null sometimes
		if (colsArrays.length == 0)
			return max;
		CTCols colsArray = colsArrays[0];
		for (int i = 0; i < colsArray.sizeOfColArray(); i++) {
			CTCol colArray = colsArray.getColArray(i);
			max = Math.max(max, (int) colArray.getMax()-1);//in CT col it is 1base, -1 to 0base.
		}
		return max;
	}
}	
