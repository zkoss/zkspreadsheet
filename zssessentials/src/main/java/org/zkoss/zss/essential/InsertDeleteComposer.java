package org.zkoss.zss.essential;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.DeleteShift;
import org.zkoss.zss.api.Range.InsertCopyOrigin;
import org.zkoss.zss.api.Range.InsertShift;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Label;


/**
 * This class demonstrates insertion and deletion API usage
 * @author dennis, Hawk
 *
 */
public class InsertDeleteComposer extends SelectorComposer<Component> {

	private static final long serialVersionUID = 1L;
	
	@Wire
	private Label selectionLabel;
	@Wire
	private Spreadsheet ss;
	
	@Listen("onCellSelection = #ss")
	public void onCellSelection(){
		selectionLabel.setValue(Ranges.getAreaReference(ss.getSelection()));
	}
	
	@Listen("onClick = #insertRow")
	public void onInsertRow(){
		Rect selection = ss.getSelection();
		Range range = Ranges.range(ss.getSelectedSheet(),selection.getRow(),selection.getColumn());

		//to affect all columns
		range = range.toRowRange(); 
		//shift existing row down and copy style from above cell 
		CellOperationUtil.insert(range, InsertShift.DOWN, InsertCopyOrigin.FORMAT_LEFT_ABOVE);
		ss.setMaxVisibleRows(ss.getMaxVisibleRows()+1);
	}
	
	@Listen("onClick = #insertRow10")
	public void onInsertRow10(){
		Rect selection = ss.getSelection();
		//mark to insert to the range that contains 10 row
		Range range = Ranges.range(ss.getSelectedSheet(), selection.getRow(),
				selection.getColumn(), selection.getRow() + 9,
				selection.getColumn());

		//to affect all columns
		range = range.toRowRange(); 
		//shift existing row down and copy style from above cell 
		CellOperationUtil.insert(range ,InsertShift.DOWN, InsertCopyOrigin.FORMAT_LEFT_ABOVE);
		ss.setMaxVisibleRows(ss.getMaxVisibleRows()+10);
	}
	
	@Listen("onClick = #deleteRows")
	public void onDeleteRows(){
		Rect selection = ss.getSelection();
		Range range = Ranges.range(ss.getSelectedSheet(), selection);

		//to affect all columns
		range = range.toRowRange(); 
		//shift existing row up 
		CellOperationUtil.delete(range, DeleteShift.UP);
	}	
	
	
	@Listen("onClick = #insertColumn")
	public void onInsertColumn(){
		Rect selection = ss.getSelection();
		Range range = Ranges.range(ss.getSelectedSheet(),selection.getRow(),selection.getColumn());

		//to affect all rows
		range = range.toColumnRange(); 
		//shift existing row right and copy style from left cell 
		CellOperationUtil.insert(range, InsertShift.RIGHT, InsertCopyOrigin.FORMAT_LEFT_ABOVE);
		ss.setMaxVisibleColumns(ss.getMaxVisibleColumns()+1);
		
	}
	
	@Listen("onClick = #insert3Columns")
	public void onInsert3Columns(){
		Rect selection = ss.getSelection();
		//select the range that contains 3 column
		Range range = Ranges.range(ss.getSelectedSheet(), selection.getRow(),
				selection.getColumn(), selection.getRow() ,
				selection.getColumn()+2);

		//select all columns
		range = range.toColumnRange(); 
		//shift existing row right and copy style from left cells 
		CellOperationUtil.insert(range, InsertShift.RIGHT, InsertCopyOrigin.FORMAT_LEFT_ABOVE);
		ss.setMaxVisibleColumns(ss.getMaxVisibleColumns()+3);
	}
	
	@Listen("onClick = #deleteColumns")
	public void onDeleteColumns(){
		Rect selection = ss.getSelection();
		//mark to insert to the range that contains 10 row
		Range range = Ranges.range(ss.getSelectedSheet(), selection);

		//to affect all rows
		range = range.toColumnRange(); 
		//shift existing row left 
		CellOperationUtil.delete(range, DeleteShift.LEFT);
	}
	
	@Listen("onClick = #insertCellDown")
	public void onInsertCellDown(){
		Rect selection = ss.getSelection();

		Range range = Ranges.range(ss.getSelectedSheet(), selection);
		
		//shift existing row down and copy style from above cell 
		CellOperationUtil.insert(range, InsertShift.DOWN, InsertCopyOrigin.FORMAT_LEFT_ABOVE);
	}
	
	@Listen("onClick = #insertCellRight")
	public void onInsertCellRight(){
		Rect selection = ss.getSelection();

		Range range = Ranges.range(ss.getSelectedSheet(), selection);
		
		//shift existing row right and copy style from left cell 
		CellOperationUtil.insert(range, InsertShift.RIGHT, InsertCopyOrigin.FORMAT_LEFT_ABOVE);
	}
	
	@Listen("onClick = #deleteCellUp")
	public void onDeleteCellUp(){
		Rect selection = ss.getSelection();

		Range range = Ranges.range(ss.getSelectedSheet(), selection);
		
		//move existing (bottom) cells up 
		CellOperationUtil.delete(range, DeleteShift.UP);
	}
	
	@Listen("onClick = #deleteCellLeft")
	public void onDeleteCellLeft(){
		Rect selection = ss.getSelection();

		Range range = Ranges.range(ss.getSelectedSheet(), selection);
		
		//move existing (right) cells left
		CellOperationUtil.delete(range, DeleteShift.LEFT);
	}
}



