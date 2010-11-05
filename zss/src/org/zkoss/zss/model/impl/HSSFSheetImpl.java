/* HSSFSheetImpl.java

	Purpose:
		
	Description:
		
	History:
		Sep 3, 2010 10:23:47 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;


import org.zkoss.poi.hssf.model.HSSFFormulaParser;
import org.zkoss.poi.hssf.model.InternalSheet;
import org.zkoss.poi.hssf.model.InternalWorkbook;
import org.zkoss.poi.hssf.record.CellValueRecordInterface;
import org.zkoss.poi.hssf.record.DVRecord;
import org.zkoss.poi.hssf.record.HyperlinkRecord;
import org.zkoss.poi.hssf.record.NameRecord;
import org.zkoss.poi.hssf.record.NoteRecord;
import org.zkoss.poi.hssf.record.Record;
import org.zkoss.poi.hssf.record.RecordBase;
import org.zkoss.poi.hssf.record.aggregates.DataValidityTable;
import org.zkoss.poi.hssf.record.aggregates.RecordAggregate.RecordVisitor;
import org.zkoss.poi.hssf.record.formula.Ptg;
import org.zkoss.poi.hssf.usermodel.HSSFCell;
import org.zkoss.poi.hssf.usermodel.HSSFCellHelper;
import org.zkoss.poi.hssf.usermodel.HSSFCellStyle;
import org.zkoss.poi.hssf.usermodel.HSSFComment;
import org.zkoss.poi.hssf.usermodel.HSSFDataValidation;
import org.zkoss.poi.hssf.usermodel.HSSFRow;
import org.zkoss.poi.hssf.usermodel.HSSFRowHelper;
import org.zkoss.poi.hssf.usermodel.HSSFSheet;
import org.zkoss.poi.hssf.usermodel.HSSFSheetHelper;
import org.zkoss.poi.hssf.usermodel.HSSFWorkbookHelper;
import org.zkoss.poi.ss.SpreadsheetVersion;
import org.zkoss.poi.ss.formula.PtgShifter;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.DataValidation;
import org.zkoss.poi.ss.usermodel.DataValidationConstraint;
import org.zkoss.poi.ss.usermodel.DataValidationHelper;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.poi.ss.usermodel.DataValidationConstraint.ValidationType;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.poi.ss.util.CellRangeAddressList;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Range;

/**
 * Implementation of {@link Sheet} based on HSSFSheet.
 * @author henrichen
 *
 */
public class HSSFSheetImpl extends HSSFSheet {
	private final HSSFSheetHelper _helper; //helper to lift the package protection
	

	//--HSSFSheet--//
	protected HSSFSheetImpl(HSSFBookImpl workbook) {
		super(workbook);
		_helper = new HSSFSheetHelper(this);
	}
	
    protected HSSFSheetImpl(HSSFBookImpl workbook, InternalSheet sheet) {
    	super(workbook, sheet);
		_helper = new HSSFSheetHelper(this);
    }
    
    //20100520, henrichen@zkoss.org: Shift rows only, don't handle formula
    /**
     * Shifts rows between startRow and endRow n number of rows.
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
     * @param copyRowHeight whether to copy the row height during the shift
     * @param resetOriginalRowHeight whether to set the original row's height to the default
     * @param moveComments whether to move comments at the same time as the cells they are attached to
     * @param clearRest whether clear the rest row after shifted endRow (meaningful only when n < 0)
     * @param copyOrigin copy format from the above/below row for the inserted rows(meaningful only when n > 0)
     * @return List of shifted merge ranges 
     */
    public List<CellRangeAddress[]> shiftRowsOnly(int startRow, int endRow, int n,
            boolean copyRowHeight, boolean resetOriginalRowHeight, boolean moveComments, boolean clearRest, int copyOrigin) {
    	//prepare source format row
    	final int srcRownum = n <= 0 ? -1 : copyOrigin == Range.FORMAT_RIGHTBELOW ? startRow : copyOrigin == Range.FORMAT_LEFTABOVE ? startRow - 1 : -1;
    	final HSSFRow srcRow = srcRownum >= 0 ? getRow(srcRownum) : null;
    	final Map<Integer, Cell> srcCells = srcRow != null ? BookHelper.copyRowCells(srcRow, srcRow.getFirstCellNum(), srcRow.getLastCellNum()) : null;
    	final short srcHeight = srcRow != null ? srcRow.getHeight() : -1;
    	final HSSFCellStyle srcStyle = srcRow != null ? srcRow.getRowStyle() : null;
    	
        int s, inc;
        if (n < 0) {
            s = startRow;
            inc = 1;
        } else {
            s = endRow;
            inc = -1;
        }
        NoteRecord[] noteRecs;
        if (moveComments) {
            noteRecs = _helper.getInternalSheet().getNoteRecords();
        } else {
            noteRecs = NoteRecord.EMPTY_ARRAY;
        }

        final int maxcol = SpreadsheetVersion.EXCEL97.getLastColumnIndex();
        final int maxrow = SpreadsheetVersion.EXCEL97.getLastRowIndex();
        final List<CellRangeAddress[]> shiftedRanges = BookHelper.shiftMergedRegion(this, startRow, 0, endRow, maxcol, n, false);
        _helper.getInternalSheet().getPageSettings().shiftRowBreaks(startRow, endRow, n);
        
        for ( int rowNum = s; rowNum >= startRow && rowNum <= endRow && rowNum >= 0 && rowNum <= maxrow; rowNum += inc ) {
            HSSFRow row = getRow( rowNum );
            // notify all cells in this row that we are going to shift them,
            // it can throw IllegalStateException if the operation is not allowed, for example,
            // if the row contains cells included in a multi-cell array formula
            if(row != null) notifyRowShifting(row);

            final int newRowNum = rowNum + n;
            final boolean inbound = newRowNum >= 0 && newRowNum <= maxrow;
            
            if (!inbound) {
            	if (row != null) {
                    if (resetOriginalRowHeight) {
                        row.setHeight((short)-1);
                    }
                    new HSSFRowHelper(row).removeAllCells();
            	}
            	continue;
            }
            
            HSSFRow row2Replace = getRow( rowNum + n );
            if (row != null) {
	            if ( row2Replace == null )
	                row2Replace = createRow( rowNum + n );
	
	            // Remove all the old cells from the row we'll
	            //  be writing too, before we start overwriting
	            //  any cells. This avoids issues with cells
	            //  changing type, and records not being correctly
	            //  overwritten
	            new HSSFRowHelper(row2Replace).removeAllCells();
            } else {
	            // If this row doesn't exist, shall also remove
	            //  the empty destination row
            	if (row2Replace != null) {
            		removeRow(row2Replace);
            	}
	            continue; // Nothing to do for this row
            }

            // Fix up row heights if required
            if (copyRowHeight) {
                row2Replace.setHeight(row.getHeight());
            }
            if (resetOriginalRowHeight) {
                row.setHeight((short)0xff);
            }

            // Copy each cell from the source row to
            //  the destination row
            for(Iterator<Cell> cells = row.cellIterator(); cells.hasNext(); ) {
                HSSFCell cell = (HSSFCell)cells.next();
                row.removeCell( cell );
                CellValueRecordInterface cellRecord = new HSSFCellHelper(cell).getCellValueRecord();
                cellRecord.setRow( rowNum + n );
                new HSSFRowHelper(row2Replace).createCellFromRecord( cellRecord );
                _helper.getInternalSheet().addValueRecord( rowNum + n, cellRecord );
            }
            // Now zap all the cells in the source row
            new HSSFRowHelper(row).removeAllCells();

            // Move comments from the source row to the
            //  destination row. Note that comments can
            //  exist for cells which are null
            if(moveComments) {
                // This code would get simpler if NoteRecords could be organised by HSSFRow.
                for(int i=noteRecs.length-1; i>=0; i--) {
                    NoteRecord nr = noteRecs[i];
                    if (nr.getRow() != rowNum) {
                        continue;
                    }
                    HSSFComment comment = getCellComment(rowNum, nr.getColumn());
                    if (comment != null) {
                       comment.setRow(rowNum + n);
                    }
                }
            }
        }
        
        //handle inserted rows
        if (srcRow != null) {
        	final int row2 = Math.min(startRow + n - 1, SpreadsheetVersion.EXCEL97.getLastRowIndex());
        	for ( int rownum = startRow; rownum <= row2; ++rownum) {
        		HSSFRow row = getRow(rownum);
        		if (row == null) {
        			row = createRow(rownum); 
        		}
        		row.setHeight(srcHeight); //height
        		if (srcStyle != null) {
        			row.setRowStyle((HSSFCellStyle)BookHelper.copyFromStyleExceptBorder(getBook(), srcStyle));//style
        		}
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
        shiftHyperlinks(startRow, endRow, n, 0, maxcol, 0);
        
        //special case1: endRow < startRow
        //special case2: (endRow - startRow + 1) < ABS(n)
        if (n < 0) {
        	if (endRow < startRow) { //special case1
	    		final int orgStartRow = startRow + n;
	            for ( int rowNum = orgStartRow; rowNum >= orgStartRow && rowNum <= endRow && rowNum >= 0 && rowNum <= maxrow; ++rowNum) {
	                final HSSFRow row = getRow( rowNum );
	                if (row != null) {
	                	removeRow(row);
	                }
	            }
	            removeHyperlinks(orgStartRow, endRow, 0, maxcol);
            } else if (clearRest) { //special case 2
            	final int orgStartRow = endRow + n + 1;
            	if (orgStartRow <= startRow) {
    	            for ( int rowNum = orgStartRow; rowNum >= orgStartRow && rowNum <= startRow && rowNum >= 0 && rowNum <= maxrow; ++rowNum) {
    	                final HSSFRow row = getRow( rowNum );
    	                if (row != null) {
    	                	removeRow(row);
    	                }
    	            }
    	            removeHyperlinks(orgStartRow, startRow, 0, maxcol);
            	}
        	}
        }
        final int lastrow = getLastRowNum();
        final int firstrow = getFirstRowNum();
        if ( endRow == lastrow || endRow + n > lastrow ) setLastRowNum(Math.min( endRow + n, maxrow ));
        if ( startRow == firstrow || startRow + n < firstrow ) setFirstRowNum(Math.max( startRow + n, 0 ));

        // Update any formulas on this sheet that point to
        //  rows which have been moved
        int sheetIndex = _workbook.getSheetIndex(this);
        short externSheetIndex = _book.checkExternSheet(sheetIndex, sheetIndex);
        PtgShifter shifter = new PtgShifter(externSheetIndex, startRow, endRow, n, 0, maxcol, 0, SpreadsheetVersion.EXCEL97);
        updateNamesAfterCellShift(shifter);
        
        return shiftedRanges;
    }
    
    //20100525, henrichen@zkoss.org: Shift columns only, don't handle formula
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
    	//prepared inserting column format
    	final int srcCol = n <= 0 ? -1 : copyOrigin == Range.FORMAT_RIGHTBELOW ? startCol : copyOrigin == Range.FORMAT_LEFTABOVE ? startCol - 1 : -1; 
    	final CellStyle colStyle = srcCol >= 0 ? getColumnStyle(srcCol) : null;
    	final int colWidth = srcCol >= 0 ? getColumnWidth(srcCol) : -1; 
    	final Map<Integer, Cell> cells = srcCol >= 0 ? new HashMap<Integer, Cell>() : null;
    	
    	int maxColNum = -1;
        for ( int rowNum = getFirstRowNum(), endRowNum = getLastRowNum(); rowNum <= endRowNum; ++rowNum ) {
            HSSFRow row = getRow( rowNum );
            
            if (row == null) continue;
            
            if (endCol < 0) {
	            final int colNum = row.getLastCellNum() - 1;
	            if (colNum > maxColNum)
	            	maxColNum = colNum;
            }
            
            if (cells != null) {
           		final Cell cell = row.getCell(srcCol);
           		if (cell != null) {
           			cells.put(new Integer(rowNum), cell);
           		}
            }
            
            row.shiftCells(startCol, endCol, n, clearRest);
        }
        
        if (endCol < 0) {
        	endCol = maxColNum;
        }
        if (n > 0) {
	        if (startCol > endCol) { //nothing to do
	        	return Collections.emptyList();
	        }
        } else {
        	if ((startCol + n) > endCol) { //nothing to do
	        	return Collections.emptyList();
        	}
        }
        
        int s, inc;
        if (n < 0) {
            s = startCol;
            inc = 1;
        } else {
            s = endCol;
            inc = -1;
        }
        
        NoteRecord[] noteRecs;
        if (moveComments) {
            noteRecs = _helper.getInternalSheet().getNoteRecords();
        } else {
            noteRecs = NoteRecord.EMPTY_ARRAY;
        }

        final int maxrow = SpreadsheetVersion.EXCEL97.getLastRowIndex();
        final int maxcol = SpreadsheetVersion.EXCEL97.getLastColumnIndex();
        final List<CellRangeAddress[]> shiftedRanges = BookHelper.shiftMergedRegion(this, 0, startCol, maxrow, endCol, n, true);
        _helper.getInternalSheet().getPageSettings().shiftColumnBreaks((short)startCol, (short)endCol, (short)n);

        // Fix up column width and comment if required
        if (moveComments || copyColWidth || resetOriginalColWidth) {
        	final int defaultColumnWidth = getDefaultColumnWidth();
	        for ( int colNum = s; colNum >= startCol && colNum <= endCol && colNum >= 0 && colNum <= maxcol; colNum += inc ) {
	        	final int newColNum = colNum + n;
		        if (copyColWidth) {
		            setColumnWidth(newColNum, getColumnWidth(colNum));
		        }
		        if (resetOriginalColWidth) {
		            setColumnWidth(colNum, defaultColumnWidth);
		        }
	            // Move comments from the source column to the
	            //  destination column. Note that comments can
	            //  exist for cells which are null
	            if(moveComments) {
	                for(int i=noteRecs.length-1; i>=0; i--) {
	                    NoteRecord nr = noteRecs[i];
	                    if (nr.getColumn() != colNum) {
	                        continue;
	                    }
	                    HSSFComment comment = getCellComment(nr.getRow(), colNum);
	                    if (comment != null) {
	                       comment.setColumn(newColNum);
	                    }
	                }
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
		            final HSSFRow row = getRow(cellEntry.getKey().intValue());
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
        shiftHyperlinks(0, maxrow, 0, startCol, endCol, n);
        
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
        int sheetIndex = _workbook.getSheetIndex(this);
        short externSheetIndex = _book.checkExternSheet(sheetIndex, sheetIndex);
        PtgShifter shifter = new PtgShifter(externSheetIndex, 0, SpreadsheetVersion.EXCEL97.getLastRowIndex(), 0, startCol, endCol, n, SpreadsheetVersion.EXCEL97);
        updateNamesAfterCellShift(shifter);
        
        return shiftedRanges;
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
    	
    	//prepared inserting column format
    	final int srcCol = n <= 0 ? -1 : copyOrigin == Range.FORMAT_RIGHTBELOW ? startCol : copyOrigin == Range.FORMAT_LEFTABOVE ? startCol - 1 : -1;
    	final CellStyle colStyle = srcCol >= 0 ? getColumnStyle(srcCol) : null;
    	final int colWidth = srcCol >= 0 ? getColumnWidth(srcCol) : -1; 
    	final Map<Integer, Cell> cells = srcCol >= 0 ? new HashMap<Integer, Cell>() : null;
    	
    	int startRow = Math.max(tRow, getFirstRowNum());
    	int endRow = Math.min(bRow, getLastRowNum());
    	int maxColNum = -1;
        for ( int rowNum = startRow, endRowNum = endRow; rowNum <= endRowNum; ++rowNum ) {
            HSSFRow row = getRow( rowNum );
            
            if (row == null) continue;
            
            if (endCol < 0) {
	            final int colNum = row.getLastCellNum() - 1;
	            if (colNum > maxColNum)
	            	maxColNum = colNum;
            }
            
            if (n > 0 && cells != null) {
           		final Cell cell = row.getCell(srcCol);
           		if (cell != null) {
           			cells.put(new Integer(rowNum), cell);
           		}
            }
            
            row.shiftCells(startCol, endCol, n, clearRest);
        }
        
        if (endCol < 0) {
        	endCol = maxColNum;
        }
        if (n > 0) {
	        if (startCol > endCol) { //nothing to do
	        	return Collections.emptyList();
	        }
        } else {
        	if ((startCol + n) > endCol) { //nothing to do
	        	return Collections.emptyList();
        	}
        }
        
        int s, inc;
        if (n < 0) {
            s = startCol;
            inc = 1;
        } else {
            s = endCol;
            inc = -1;
        }
        
        NoteRecord[] noteRecs;
        if (moveComments) {
            noteRecs = _helper.getInternalSheet().getNoteRecords();
        } else {
            noteRecs = NoteRecord.EMPTY_ARRAY;
        }

        final int maxrow = SpreadsheetVersion.EXCEL97.getLastRowIndex();
        final int maxcol = SpreadsheetVersion.EXCEL97.getLastColumnIndex();
        final List<CellRangeAddress[]> shiftedRanges = BookHelper.shiftMergedRegion(this, tRow, startCol, bRow, endCol, n, true);
        final boolean wholeColumn = tRow == 0 && bRow == maxrow; 
        if (wholeColumn) { 
        	_helper.getInternalSheet().getPageSettings().shiftColumnBreaks((short)startCol, (short)endCol, (short)n);
        }

        // Fix up column width and comment if required
        if (moveComments || copyColWidth || resetOriginalColWidth) {
        	final int defaultColumnWidth = getDefaultColumnWidth();
	        for ( int colNum = s; colNum >= startCol && colNum <= endCol && colNum >= 0 && colNum <= maxcol; colNum += inc ) {
	        	final int newColNum = colNum + n;
	        	if (wholeColumn) {
			        if (copyColWidth) {
			            setColumnWidth(newColNum, getColumnWidth(colNum));
			        }
			        if (resetOriginalColWidth) {
			            setColumnWidth(colNum, defaultColumnWidth);
			        }
	        	}
	            // Move comments from the source column to the
	            //  destination column. Note that comments can
	            //  exist for cells which are null
	            if(moveComments) {
	                for(int i=noteRecs.length-1; i>=0; i--) {
	                    NoteRecord nr = noteRecs[i];
	                    if (nr.getColumn() != colNum || nr.getRow() < tRow || nr.getRow() > bRow) { //not in range
	                        continue;
	                    }
	                    HSSFComment comment = getCellComment(nr.getRow(), colNum);
	                    if (comment != null) {
	                       comment.setColumn(newColNum);
	                    }
	                }
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
	        		setDefaultColumnStyle(col, BookHelper.copyFromStyleExceptBorder(getBook(), colStyle));
	        	}
	        }
	        if (cells != null) {
		        for (Entry<Integer, Cell> cellEntry : cells.entrySet()) {
		            final HSSFRow row = getRow(cellEntry.getKey().intValue());
		            final Cell srcCell = cellEntry.getValue();
		            final CellStyle srcStyle = srcCell.getCellStyle();
		        	for (int col = startCol; col <= col2 ; ++col) {
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
        shiftHyperlinks(tRow, bRow, 0, startCol, endCol, n);
        
        //special case1: endCol < startCol
        //special case2: (endCol - startCol + 1) < ABS(n)
        if (n < 0) {
        	if (endCol < startCol) { //special case1
	    		final int replacedStartCol = startCol + n;
	    		removeHyperlinks(tRow, bRow, replacedStartCol, endCol);
            } else if (clearRest) { //special case 2
            	final int replacedStartCol = endCol + n + 1;
            	if (replacedStartCol <= startCol) {
    	    		removeHyperlinks(tRow, bRow, replacedStartCol, startCol);
            	}
        	}
        }

        
        // Update any formulas on this sheet that point to
        // columns which have been moved
        int sheetIndex = _workbook.getSheetIndex(this);
        short externSheetIndex = _book.checkExternSheet(sheetIndex, sheetIndex);
        PtgShifter shifter = new PtgShifter(externSheetIndex, tRow, bRow, 0, startCol, endCol, n, SpreadsheetVersion.EXCEL97);
        updateNamesAfterCellShift(shifter);
        
        return shiftedRanges;
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
    	//prepare source format row
    	final int srcRownum = n <= 0 ? -1 : copyOrigin == Range.FORMAT_RIGHTBELOW ? startRow : copyOrigin == Range.FORMAT_LEFTABOVE ? startRow - 1 : -1;
    	final HSSFRow srcRow = srcRownum >= 0 ? getRow(srcRownum) : null;
    	final Map<Integer, Cell> srcCells = srcRow != null ? BookHelper.copyRowCells(srcRow, lCol, rCol) : null;
    	final short srcHeight = srcRow != null ? srcRow.getHeight() : -1;
    	final HSSFCellStyle srcStyle = srcRow != null ? srcRow.getRowStyle() : null;
    	
        final int maxrow = SpreadsheetVersion.EXCEL97.getLastRowIndex();
        if (endRow < 0) {
        	endRow = maxrow;
        }
        if (n > 0) {
	        if (startRow > endRow ) { //nothing to do
	        	return Collections.emptyList();
	        }
        } else {
        	if ((startRow + n) > endRow) { //nothing to do
	        	return Collections.emptyList();
        	}
        }
        
        int s, inc;
        if (n < 0) {
            s = startRow;
            inc = 1;
        } else {
            s = endRow;
            inc = -1;
        }
        NoteRecord[] noteRecs;
        if (moveComments) {
            noteRecs = _helper.getInternalSheet().getNoteRecords();
        } else {
            noteRecs = NoteRecord.EMPTY_ARRAY;
        }

        final int maxcol = SpreadsheetVersion.EXCEL97.getLastColumnIndex();
        final List<CellRangeAddress[]> shiftedRanges = BookHelper.shiftMergedRegion(this, startRow, lCol, endRow, rCol, n, false);
        final boolean wholeRow = lCol == 0 && rCol == maxcol; 
        if (wholeRow) { 
        	_helper.getInternalSheet().getPageSettings().shiftRowBreaks(startRow, endRow, n);
        }
        for ( int rowNum = s; rowNum >= startRow && rowNum <= endRow && rowNum >= 0 && rowNum <= maxrow; rowNum += inc ) {
            HSSFRow row = getRow( rowNum );
            // notify all cells in this row that we are going to shift them,
            // it can throw IllegalStateException if the operation is not allowed, for example,
            // if the row contains cells included in a multi-cell array formula
            if(row != null) notifyRowShifting(row); //TODO: add lCol, rCol information

            final int newRowNum = rowNum + n;
            final boolean inbound = newRowNum >= 0 && newRowNum <= maxrow;
            
            if (!inbound) {
            	if (row != null) {
            		if (wholeRow) {
	                    if (resetOriginalRowHeight) {
	                        row.setHeight((short)-1); //default height
	                    }
	                    new HSSFRowHelper(row).removeAllCells();
            		} else {
            			removeCells(row, lCol, rCol);
            		}
            	}
            	continue;
            }
            
            HSSFRow row2Replace = getRow( rowNum + n );
            if (row != null) {
	            if ( row2Replace == null )
	                row2Replace = createRow( rowNum + n );
	
	            // Remove all the old cells from the row we'll
	            //  be writing too, before we start overwriting
	            //  any cells. This avoids issues with cells
	            //  changing type, and records not being correctly
	            //  overwritten
	            if (wholeRow) {
	            	new HSSFRowHelper(row2Replace).removeAllCells();
	            } else {
	            	removeCells(row2Replace, lCol, rCol);
	            }
            } else {
	            // If this row doesn't exist, shall also remove
	            //  the empty destination row
            	if (row2Replace != null) {
            		if (wholeRow) {
            			removeRow(row2Replace);
            		} else {
    	            	removeCells(row2Replace, lCol, rCol);
            		}
            	}
	            continue; // Nothing to do for this row
            }

            // Fix up row heights if required
            if (wholeRow) {
	            if (copyRowHeight) {
	                row2Replace.setHeight(row.getHeight());
	            }
	            if (resetOriginalRowHeight) {
	                row.setHeight((short)0xff);
	            }
            }

            // Copy each cell from the source row to
            //  the destination row
            if (wholeRow) {
	            for(Iterator<Cell> cells = row.cellIterator(); cells.hasNext(); ) {
	                HSSFCell cell = (HSSFCell)cells.next();
	                row.removeCell( cell );
	                CellValueRecordInterface cellRecord = new HSSFCellHelper(cell).getCellValueRecord();
	                cellRecord.setRow( rowNum + n );
	                new HSSFRowHelper(row2Replace).createCellFromRecord( cellRecord );
	                _helper.getInternalSheet().addValueRecord( rowNum + n, cellRecord );
	            }
	            // Now zap all the cells in the source row
	            new HSSFRowHelper(row).removeAllCells();
            } else {
            	final int startCol = Math.max(row.getFirstCellNum(), lCol);
            	final int endCol = Math.min(row.getLastCellNum(), rCol);
	            for(int col = startCol; col <= endCol; ++col) {
	                HSSFCell cell = (HSSFCell)row.getCell(col);
	                if (cell == null) {
	                	continue;
	                }
	                row.removeCell( cell );
	                CellValueRecordInterface cellRecord = new HSSFCellHelper(cell).getCellValueRecord();
	                cellRecord.setRow( rowNum + n );
	                new HSSFRowHelper(row2Replace).createCellFromRecord( cellRecord );
	                _helper.getInternalSheet().addValueRecord( rowNum + n, cellRecord );
	            }
	            // Now zap the cells in the source row
	            removeCells(row, lCol, rCol);
            }

            // Move comments from the source row to the
            //  destination row. Note that comments can
            //  exist for cells which are null
            if(moveComments) {
                // This code would get simpler if NoteRecords could be organised by HSSFRow.
                for(int i=noteRecs.length-1; i>=0; i--) {
                    NoteRecord nr = noteRecs[i];
                    if (nr.getRow() != rowNum || nr.getColumn() < lCol || nr.getColumn() > rCol) {
                        continue;
                    }
                    HSSFComment comment = getCellComment(rowNum, nr.getColumn());
                    if (comment != null) {
                       comment.setRow(rowNum + n);
                    }
                }
            }
        }
        
        //handle inserted rows
        if (srcRow != null) {
        	final int row2 = Math.min(startRow + n - 1, SpreadsheetVersion.EXCEL97.getLastRowIndex());
        	for(int rownum = startRow; rownum <= row2; ++rownum) {
        		HSSFRow row = getRow(rownum);
        		if (row == null) {
        			row = createRow(rownum); 
        		}
        		if (wholeRow) {
        			row.setHeight(srcHeight);
        			row.setRowStyle((HSSFCellStyle)BookHelper.copyFromStyleExceptBorder(getBook(), srcStyle));
        		}
                if (srcCells != null) {
	        		for(Entry<Integer, Cell> cellEntry : srcCells.entrySet()) {
	        			final int colnum = cellEntry.getKey().intValue();
	        			final Cell srcCell = cellEntry.getValue();
	        			final CellStyle cellStyle = srcCell.getCellStyle(); 
	        			Cell dstCell = row.getCell(colnum);
	        			if (dstCell == null) {
	        				dstCell = row.createCell(colnum);
	        			}
	        			dstCell.setCellStyle(BookHelper.copyFromStyleExceptBorder(getBook(), cellStyle));
	        		}
                }
        	}
        }
        
        // Shift Hyperlinks which have been moved
        shiftHyperlinks(startRow, endRow, n, lCol, rCol, 0);
        
        //special case1: endRow < startRow
        //special case2: (endRow - startRow + 1) < ABS(n)
        if (n < 0) {
        	if (endRow < startRow) { //special case1
	    		final int orgStartRow = startRow + n;
	            for ( int rowNum = orgStartRow; rowNum >= orgStartRow && rowNum <= endRow && rowNum >= 0 && rowNum <= maxrow; ++rowNum) {
	                final HSSFRow row = getRow( rowNum );
	                if (row != null) {
	                	if (wholeRow) {
	                		removeRow(row);
	                	} else {
	                		removeCells(row, lCol, rCol);
	                	}
	                }
	            }
	            removeHyperlinks(orgStartRow, endRow, lCol, rCol);
            } else if (clearRest) { //special case 2
            	final int orgStartRow = endRow + n + 1;
            	if (orgStartRow <= startRow) {
    	            for ( int rowNum = orgStartRow; rowNum >= orgStartRow && rowNum <= startRow && rowNum >= 0 && rowNum <= maxrow; ++rowNum) {
    	                final HSSFRow row = getRow( rowNum );
    	                if (row != null) {
    	                	if (wholeRow) {
    	                		removeRow(row);
    	                	} else {
    	                		removeCells(row, lCol, rCol);
    	                	}
    	                }
    	            }
    	            removeHyperlinks(orgStartRow, startRow, lCol, rCol);
            	}
        	}
        }
        final int lastrow = getLastRowNum();
        final int firstrow = getFirstRowNum();
        if ( endRow == lastrow || endRow + n > lastrow ) setLastRowNum(Math.min( endRow + n, maxrow));
        if ( startRow == firstrow || startRow + n < firstrow ) setFirstRowNum(Math.max( startRow + n, 0 ));

        // Update any formulas on this sheet that point to
        //  rows which have been moved
        int sheetIndex = _workbook.getSheetIndex(this);
        short externSheetIndex = _book.checkExternSheet(sheetIndex, sheetIndex);
        PtgShifter shifter = new PtgShifter(externSheetIndex, startRow, endRow, n, lCol, rCol, 0, SpreadsheetVersion.EXCEL97);
        updateNamesAfterCellShift(shifter);
        
        return shiftedRanges;
    }
    
    //20100701, henrichen@zkoss.org: remove cells of the row between specified left column and right column
    private void removeCells(Row row, int lCol, int rCol) {
    	final int startCol = Math.max(row.getFirstCellNum(), lCol);
    	final int endCol = Math.min(row.getLastCellNum(), rCol);
		for(int col = startCol; col <= endCol; ++col) {
			final Cell cell = row.getCell(col);
			if (cell != null) {
				row.removeCell(cell);
			}
		}
    }

    public List<CellRangeAddress[]> shiftBothRange(int tRow, int bRow, int nRow, int lCol, int rCol, int nCol,
        boolean moveComments) {
        if (nRow > 0) {
	        if (tRow > bRow ) { //nothing to do
	        	return Collections.emptyList();
	        }
        } else {
        	if ((tRow + nRow) > bRow) { //nothing to do
	        	return Collections.emptyList();
        	}
        }
        
        int s, inc;
        if (nRow < 0) {
            s = tRow;
            inc = 1;
        } else {
            s = bRow;
            inc = -1;
        }
        NoteRecord[] noteRecs;
        if (moveComments) {
            noteRecs = _helper.getInternalSheet().getNoteRecords();
        } else {
            noteRecs = NoteRecord.EMPTY_ARRAY;
        }

        final int maxrow = SpreadsheetVersion.EXCEL97.getLastRowIndex();
        final int maxcol = SpreadsheetVersion.EXCEL97.getLastColumnIndex();
        final List<CellRangeAddress[]> shiftedRanges = BookHelper.shiftBothMergedRegion(this, tRow, lCol, bRow, rCol, nRow, nCol);
        for ( int rowNum = s; rowNum >= tRow && rowNum <= bRow && rowNum >= 0 && rowNum <= maxrow; rowNum += inc ) {
            HSSFRow row = getRow( rowNum );
            // notify all cells in this row that we are going to shift them,
            // it can throw IllegalStateException if the operation is not allowed, for example,
            // if the row contains cells included in a multi-cell array formula
            if(row != null) notifyRowShifting(row); //TODO: add lCol, rCol information

            final int newRowNum = rowNum + nRow;
            final boolean rowInbound = newRowNum >= 0 && newRowNum <= maxrow;
            if (!rowInbound) {
            	if (row != null) {
           			removeCells(row, lCol, rCol);
            	}
            	continue;
            }

            final int dstlCol = Math.max(lCol + nCol, 0);
            final int dstrCol = Math.min(rCol + nCol, maxcol);
            final boolean colInbound = dstlCol > maxcol || dstrCol < 0;
            
            HSSFRow row2Replace = getRow( newRowNum );
            if (row != null) {
	            if ( row2Replace == null )
	                row2Replace = createRow( newRowNum );
	
	            // Remove all the old cells from the row we'll
	            //  be writing too, before we start overwriting
	            //  any cells. This avoids issues with cells
	            //  changing type, and records not being correctly
	            //  overwritten
	            if (colInbound) {
	            	removeCells(row2Replace, dstlCol, dstrCol);
	            }
            } else {
	            // If this row doesn't exist, shall also remove
	            //  the empty destination row
            	if (row2Replace != null) {
            		if (colInbound) {
            			removeCells(row2Replace, dstlCol, dstrCol);
            		}
            	}
	            continue; // Nothing to do for this row
            }

            // Copy each cell from the source row to
            //  the destination row
        	final int startCol = Math.max(row.getFirstCellNum(), lCol);
        	final int endCol = Math.min(row.getLastCellNum(), rCol);
            for(int col = startCol; col <= endCol; ++col) {
                HSSFCell cell = (HSSFCell)row.getCell(col);
                if (cell == null) {
                	continue;
                }
                row.removeCell( cell );
                CellValueRecordInterface cellRecord = new HSSFCellHelper(cell).getCellValueRecord();
                cellRecord.setRow( rowNum + nRow );
                int colNum = cellRecord.getColumn() + nCol;
                if (colNum >= 0 && colNum <= maxcol) { //in bound
                	cellRecord.setColumn((short) colNum);
                }
                new HSSFRowHelper(row2Replace).createCellFromRecord( cellRecord );
                _helper.getInternalSheet().addValueRecord( rowNum + nRow, cellRecord );
            }
            // Now zap the cells in the source row
            removeCells(row, lCol, rCol);

            // Move comments from the source row to the
            //  destination row. Note that comments can
            //  exist for cells which are null
            if(moveComments) {
                // This code would get simpler if NoteRecords could be organised by HSSFRow.
                for(int i=noteRecs.length-1; i>=0; i--) {
                    NoteRecord nr = noteRecs[i];
                    if (nr.getRow() != rowNum || nr.getColumn() < lCol || nr.getColumn() > rCol) {
                        continue;
                    }
                    HSSFComment comment = getCellComment(rowNum, nr.getColumn());
                    if (comment != null) {
                    	comment.setRow(rowNum + nRow);
                    	int colNum = nr.getColumn()+nCol;
                    	if (colNum >= 0 && colNum <= maxcol) {
                    		comment.setColumn(colNum);
                    	}
                    }
                }
            }
        }
        final int lastrow = getLastRowNum();
        final int firstrow = getFirstRowNum();
        if ( bRow == lastrow || bRow + nRow > lastrow ) setLastRowNum(Math.min( bRow + nRow, maxrow));
        if ( tRow == firstrow || tRow + nRow < firstrow ) setFirstRowNum(Math.max( tRow + nRow, 0 ));

        // Shift Hyperlinks which have been moved
        shiftHyperlinks(tRow, bRow, nRow, lCol, rCol, nCol);
        
        // Update any formulas on this sheet that point to
        //  rows which have been moved
        int sheetIndex = _workbook.getSheetIndex(this);
        short externSheetIndex = _book.checkExternSheet(sheetIndex, sheetIndex);
        PtgShifter shifter = new PtgShifter(externSheetIndex, tRow, bRow, nRow, lCol, rCol, nCol, SpreadsheetVersion.EXCEL97);
        updateNamesAfterCellShift(shifter);
        
        return shiftedRanges;
    }
    
    /**
     * Updates named ranges due to moving of cells
     */
    //20100705, henrichen@zkoss.org: update named references
    private void updateNamesAfterCellShift(PtgShifter shifter) {
    	InternalWorkbook book = new HSSFWorkbookHelper(getWorkbook()).getInternalWorkbook();
        for (int i = 0 ; i < book.getNumNames() ; ++i){
            NameRecord nr = book.getNameRecord(i);
            Ptg[] ptgs = nr.getNameDefinition();
            if (shifter.adjustFormula(ptgs, nr.getSheetNumber())) {
                nr.setNameDefinition(ptgs);
            }
        }
    }
    
    //20100720, henrichen@zkoss.org: shift Hyperlinks
    /**
     * Shift Hyperlink of the specified range.
     */
    private void shiftHyperlinks(int tRow, int bRow, int nRow, int lCol, int rCol, int nCol) {
    	final int maxcol = SpreadsheetVersion.EXCEL97.getLastColumnIndex();
    	final int maxrow = SpreadsheetVersion.EXCEL97.getLastRowIndex();
        for (Iterator<RecordBase> it = _helper.getInternalSheet().getRecords().iterator(); it.hasNext(); ) {
            RecordBase rec = it.next();
            if (rec instanceof HyperlinkRecord){
                final HyperlinkRecord link = (HyperlinkRecord)rec;
                final int col = link.getFirstColumn();
                final int row = link.getFirstRow();
                if (inRange(row, tRow, bRow) && inRange(col, lCol, rCol)) {
                	final int dstrow = row + nRow;
                	final int dstcol = col + nCol;
                	if (inRange(dstrow, 0, maxrow) && inRange(dstcol, 0, maxcol)) {
                		link.setFirstColumn(dstcol);
                		link.setFirstRow(dstrow);
                		link.setLastColumn(dstcol);
                		link.setLastRow(dstrow);
                	} else {
                		it.remove();
                	}
                }
            }
        }
    }
    
    private void removeHyperlinks(int tRow, int bRow, int lCol, int rCol) {
        for (Iterator<RecordBase> it = _helper.getInternalSheet().getRecords().iterator(); it.hasNext(); ) {
            RecordBase rec = it.next();
            if (rec instanceof HyperlinkRecord){
                final HyperlinkRecord link = (HyperlinkRecord)rec;
                final int col = link.getFirstColumn();
                final int row = link.getFirstRow();
                if (inRange(row, tRow, bRow) && inRange(col, lCol, rCol)) {
               		it.remove();
                }
            }
        }
    }
    
    private final boolean inRange(int val, int min, int max) {
    	return min <= val && val <= max;
    }
    
    //20100727, henrichen@zkoss.org: handle DataValidation
	private List<DataValidation> _dataValidations;
    public List<DataValidation> getDataValidations() {
    	if (_dataValidations == null) {
    		//populate the _dataValidations list
	    	final DataValidityTable tb = _helper.getInternalSheet().getOrCreateDataValidityTable();
	    	_dataValidations = new ArrayList<DataValidation>(); 
	    	final RecordVisitor rv = new DVRecordVisitor(); 
	    	tb.visitContainedRecords(rv); 
    	}
    	return _dataValidations;
    }
    
    //20100910, henrichen@zkoss.org: nullify the cached data validation list 
    public void nullifyDataValidations() {
    	_dataValidations = null;
    }
    
    private class DVRecordVisitor implements RecordVisitor {
    	private DataValidationHelper _helper;
    	
    	private DVRecordVisitor() {
    		_helper = getDataValidationHelper();
    	}
    	
		@Override
		public void visitRecord(Record r) {
			if (r instanceof DVRecord) {
				final DVRecord dvRecord = (DVRecord) r;
				final CellRangeAddressList regions = dvRecord.getCellRangeAddress();
				final DataValidationConstraint constraint = createContraint(dvRecord);
				if (constraint != null) {
					HSSFDataValidation dataValidation = new HSSFDataValidation(regions, constraint);
					final boolean allowed = dvRecord.getEmptyCellAllowed();
					final int errStyle = dvRecord.getErrorStyle();
					final boolean showErr = dvRecord.getShowErrorOnInvalidValue();
					final boolean showPrompt = dvRecord.getShowPromptOnCellSelected();
					final boolean suppress = dvRecord.getSuppressDropdownArrow();
					
					final String promptTitle = dvRecord.getPromptTitle();
					final String promptText = dvRecord.getPromptText();
					final String errorTitle = dvRecord.getErrorTitle();
					final String errorText = dvRecord.getErrorText();
					if (showPrompt) {
						dataValidation.createPromptBox(promptTitle, promptText);
					}
					if (showErr) {
						dataValidation.createErrorBox(errorTitle, errorText);
					}
					dataValidation.setEmptyCellAllowed(allowed);
					dataValidation.setErrorStyle(errStyle);
					dataValidation.setShowErrorBox(showErr);
					dataValidation.setShowPromptBox(showPrompt);
					dataValidation.setSuppressDropDownArrow(suppress);
					
					_dataValidations.add(dataValidation);
				}
			}
		}
		
		private DataValidationConstraint createContraint(DVRecord dvRecord) {
			final int operatorType = dvRecord.getConditionOperator();
			final int validationType = dvRecord.getDataType();
			final boolean isExplicitValues = !dvRecord.getListExplicitFormula(); //whether list is given by excplicit values?
			final Ptg[] formulaPtgs1 = dvRecord.getFormula1();
			final Ptg[] formulaPtgs2 = dvRecord.getFormula2();
			String formula1 = HSSFFormulaParser.toFormulaString(_workbook, formulaPtgs1);
			//bug #64: pasteSpecial copy Validation cause Null pointer exception (Ptgs must not be null)   	
			String formula2 = formulaPtgs2 != null && formulaPtgs2.length > 0 ? HSSFFormulaParser.toFormulaString(_workbook, formulaPtgs2) : null; 
			switch(validationType) {
				case ValidationType.ANY:
					return _helper.createNumericConstraint(validationType, operatorType, formula1, formula2);
					
				case ValidationType.DECIMAL:
					return _helper.createDecimalConstraint(operatorType, formula1, formula2);
					
				case ValidationType.INTEGER:
					return _helper.createIntegerConstraint(operatorType, formula1, formula2);
					
				case ValidationType.TEXT_LENGTH:
					return _helper.createTextLengthConstraint(operatorType, formula1, formula2);
					
				case ValidationType.DATE:
					return _helper.createDateConstraint(operatorType, formula1, formula2, null /*dateFormat*/);
					
				case ValidationType.FORMULA:
					return _helper.createCustomConstraint(formula1);
					
				case ValidationType.LIST:
					if (isExplicitValues) {
						final String[] listOfValues = formula1.split("\0");
						return _helper.createExplicitListConstraint(listOfValues);
					}
					return _helper.createFormulaListConstraint(formula1);
					
				case ValidationType.TIME:
					return _helper.createTimeConstraint(operatorType, formula1, formula2);
			}
			
			return null;
		}
    }
    
    public boolean isFreezePanes() {
    	return _helper.getInternalSheet().getWindowTwo().getFreezePanes();
    }
    
    public Book getBook() {
    	return (Book) getWorkbook();
    }
    
    
}
