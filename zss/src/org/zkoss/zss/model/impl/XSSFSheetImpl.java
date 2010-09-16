/* XSSFSheetImpl.java

	Purpose:
		
	Description:
		
	History:
		Sep 3, 2010 11:07:42 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.SortedMap;
import java.util.TreeMap;
import java.util.Map.Entry;

import org.apache.poi.hssf.record.NoteRecord;
import org.apache.poi.hssf.record.formula.FormulaShifter;
import org.apache.poi.hssf.record.formula.Ptg;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFComment;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaRenderer;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.PtgShifter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.model.CommentsTable;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFRowHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.helpers.XSSFRowShifter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTComment;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCommentList;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Range;

/**
 * Implementation of {@link Sheet} based on XSSFSheet.
 * @author henrichen
 *
 */
public class XSSFSheetImpl extends XSSFSheet {
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

	public Book getBook() {
		return (Book) getWorkbook();
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
    public List<CellRangeAddress[]>  shiftRowsOnly(int startRow, int endRow, int n, boolean copyRowHeight, boolean resetOriginalRowHeight,
    		boolean moveComments, boolean clearRest, int copyOrigin) {
    	//prepare source format row
    	final int srcRownum = n <= 0 ? -1 : copyOrigin == Range.FORMAT_RIGHTBELOW ? startRow : copyOrigin == Range.FORMAT_LEFTABOVE ? startRow - 1 : -1;
    	final XSSFRow srcRow = srcRownum >= 0 ? getRow(srcRownum) : null;
    	final Map<Integer, Cell> srcCells = srcRow != null ? BookHelper.copyRowCells(srcRow, srcRow.getFirstCellNum(), srcRow.getLastCellNum()) : null;
    	final short srcHeight = srcRow != null ? srcRow.getHeight() : -1;
//    	final XSSFCellStyle srcStyle = srcRow != null ? srcRow.getRowStyle() : null;
    	
        final int maxrow = SpreadsheetVersion.EXCEL2007.getLastRowIndex();
        final int maxcol = SpreadsheetVersion.EXCEL2007.getLastColumnIndex();
        final List<CellRangeAddress[]> shiftedRanges = BookHelper.shiftMergedRegion(this, startRow, 0, endRow, maxcol, n, false);
    	
    	//shift the rows (actually change the row number only)
        for (Iterator<Row> it = rowIterator() ; it.hasNext() ; ) {
            XSSFRow row = (XSSFRow)it.next();
            int rownum = row.getRowNum();
            if(rownum < startRow) continue;
            
            final int newrownum = rownum + n;
            final boolean inbound = 0 <= newrownum && newrownum <= maxrow;
            if (!inbound) {
            	row.removeAllCells();
            	it.remove();
            	continue;
            }
            
            if (!copyRowHeight) {
                row.setHeight((short)-1);
            }

            if (canRemoveRow(startRow, endRow, n, rownum)) {
                it.remove();
            } else if (rownum >= startRow && rownum <= endRow) {
                new XSSFRowHelper(row).shift(n);
            }
            if (moveComments) {
	            final CommentsTable sheetComments = getCommentsTable(false);
	            if(sheetComments != null){
	                //TODO shift Note's anchor in the associated /xl/drawing/vmlDrawings#.vml
	                CTCommentList lst = sheetComments.getCTComments().getCommentList();
	                for (CTComment comment : lst.getCommentArray()) {
	                    CellReference ref = new CellReference(comment.getRef());
	                    if(ref.getRow() == rownum){
	                        ref = new CellReference(rownum + n, ref.getCol());
	                        comment.setRef(ref.formatAsString());
	                    }
	                }
	            }
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
        shiftHyperlinks(startRow, endRow, n, 0, maxcol, 0);
        
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
        final XSSFWorkbook wb = getWorkbook();
        int sheetIndex = wb.getSheetIndex(this);
        final PtgShifter shifter = new PtgShifter(sheetIndex, startRow, endRow, n, 0, maxcol, 0, SpreadsheetVersion.EXCEL2007);
        updateNamedRanges(wb, shifter);

        //rebuild the _rows map
        TreeMap<Integer, XSSFRow> map = new TreeMap<Integer, XSSFRow>();
        TreeMap<Integer, XSSFRow> rows = getRows();
        for(XSSFRow r : rows.values()) {
            map.put(r.getRowNum(), r);
        }
        setRows(map);
        
        return shiftedRanges;
    }
    
    private void updateNamedRanges(XSSFWorkbook wb, PtgShifter shifter) {
        XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.create(wb);
        for (int i = 0; i < wb.getNumberOfNames(); i++) {
            XSSFName name = wb.getNameAt(i);
            String formula = name.getRefersToFormula();
            int sheetIndex = name.getSheetIndex();

            Ptg[] ptgs = FormulaParser.parse(formula, fpb, FormulaType.NAMEDRANGE, sheetIndex);
            if (shifter.adjustFormula(ptgs, sheetIndex)) {
                String shiftedFmla = FormulaRenderer.toFormulaString(fpb, ptgs);
                name.setRefersToFormula(shiftedFmla);
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
    private void shiftHyperlinks(int tRow, int bRow, int nRow, int lCol, int rCol, int nCol) {
    	//TODO shift hyperlinks of the specified Range.
    }
    private void removeHyperlinks(int tRow, int bRow, int lCol, int rCol) {
    	//TODO remove hyperlinks within the Range
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
    	final XSSFRow srcRow = srcRownum >= 0 ? getRow(srcRownum) : null;
    	final Map<Integer, Cell> srcCells = srcRow != null ? BookHelper.copyRowCells(srcRow, lCol, rCol) : null;
    	final short srcHeight = srcRow != null ? srcRow.getHeight() : -1;
//    	final XSSFCellStyle srcStyle = srcRow != null ? srcRow.getRowStyle() : null;
    	
        final int maxrow = SpreadsheetVersion.EXCEL2007.getLastRowIndex();
        final int maxcol = SpreadsheetVersion.EXCEL2007.getLastColumnIndex();
        if (endRow < 0) {
        	endRow = maxrow;
        }
        final List<CellRangeAddress[]> shiftedRanges = BookHelper.shiftMergedRegion(this, startRow, lCol, endRow, rCol, n, false);
        final boolean wholeRow = lCol == 0 && rCol == maxcol; 
    	
        final List<int[]> removePairs = new ArrayList<int[]>(); //row spans to be removed 
        final TreeMap<Integer, TreeMap<Integer, XSSFCell>> rowCells = new TreeMap<Integer, TreeMap<Integer, XSSFCell>>();
        int expectRownum = startRow; //handle sparse rows which might override destination row
        for (Iterator<Row> it = rowIterator() ; it.hasNext() ; ) { //TODO use submap between startRow and endRow
            XSSFRow row = (XSSFRow)it.next();
            int rownum = row.getRowNum();
            if (rownum < startRow) continue;
            if (rownum > endRow) break; //no more
            
            final int newRownum = rownum + n;
            if (rownum > expectRownum) { //sparse row between expectRownum(inclusive) and current row(exclusive), to be removed
            	addRemovePair(removePairs, expectRownum + n, newRownum);
            }
            expectRownum = rownum + 1;
            
            final boolean inbound = 0 <= newRownum && newRownum <= maxrow;
            if (!inbound) {
            	row.removeAllCells();
            	it.remove();
            	continue;
            }
            
            if (wholeRow) {
	            if (!copyRowHeight) {
	                row.setHeight((short)-1);
	            }
	            if (canRemoveRow(startRow, endRow, n, rownum)) {
	                it.remove();
	            }
	            else if (rownum >= startRow && rownum <= endRow) {
	                new XSSFRowHelper(row).shift(n);
	            }
            } else {
            	SortedMap<Integer, XSSFCell> oldCells = row.getCells().subMap(new Integer(lCol), new Integer(rCol+1));
            	if (!oldCells.isEmpty()) {
            		TreeMap<Integer, XSSFCell> cells = new TreeMap<Integer, XSSFCell>(oldCells);
            		rowCells.put(newRownum, cells);
            	}
            }
            
            if (moveComments) {
	            final CommentsTable sheetComments = getCommentsTable(false);
	            if(sheetComments != null){
	                //TODO shift Note's anchor in the associated /xl/drawing/vmlDrawings#.vml
	                CTCommentList lst = sheetComments.getCTComments().getCommentList();
	                for (CTComment comment : lst.getCommentArray()) {
	                    CellReference ref = new CellReference(comment.getRef());
	                    final int colnum = ref.getCol();
	                    if(ref.getRow() == rownum && lCol <= colnum && colnum <= rCol){
	                        ref = new CellReference(rownum + n, colnum);
	                        comment.setRef(ref.formatAsString());
	                    }
	                }
	            }
            }
        }

        //sparse row between expectRownum(inclusive) to endRow+1(exclusive), to be removed
    	addRemovePair(removePairs, expectRownum + n, endRow+1+n);
    	
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
        shiftHyperlinks(startRow, endRow, n, 0, maxcol, 0);
        
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
        final XSSFWorkbook wb = getWorkbook();
        int sheetIndex = wb.getSheetIndex(this);
        final PtgShifter shifter = new PtgShifter(sheetIndex, startRow, endRow, n, 0, maxcol, 0, SpreadsheetVersion.EXCEL2007);
        updateNamedRanges(wb, shifter);

        //rebuild the _rows map
        if (wholeRow) {
	        TreeMap<Integer, XSSFRow> map = new TreeMap<Integer, XSSFRow>();
	        TreeMap<Integer, XSSFRow> rows = getRows();
	        for(XSSFRow r : rows.values()) {
	            map.put(r.getRowNum(), r);
	        }
	        setRows(map);
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
}
