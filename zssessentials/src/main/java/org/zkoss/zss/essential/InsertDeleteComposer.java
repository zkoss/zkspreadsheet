package org.zkoss.zss.essential;

import java.util.Arrays;
import java.util.List;

import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.DeleteShift;
import org.zkoss.zss.api.Range.InsertCopyOrigin;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.Range.InsertShift;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zul.Label;


/**
 * This class shows all the public ZK Spreadsheet you can listen to
 * @author dennis
 *
 */
public class InsertDeleteComposer extends AbstractDemoComposer {

	private static final long serialVersionUID = 1L;
	
	@Wire
	Label selectionLab;
	
	protected List<String> contirbuteAvailableBooks(){
		return Arrays.asList("full.xlsx");
	}
	
	
	@Listen("onCellSelection = #ss")
	public void onCellSelection(){
		selectionLab.setValue(Ranges.getAreaReference(ss.getSelection()));
	}
	
	@Listen("onClick = #insertRow")
	public void onInsertRow(){
		Rect selection = ss.getSelection();
		Range range = Ranges.range(ss.getSelectedSheet(),selection.getRow(),selection.getColumn());

		//to effect all columns
		range = range.toRowRange(); 
		//shift exist row down and copy style from above cell 
		range.insert(InsertShift.DOWN, InsertCopyOrigin.FORMAT_LEFT_ABOVE);
		
	}
	
	@Listen("onClick = #insertRow10")
	public void onInsertRow10(){
		Rect selection = ss.getSelection();
		//mark to insert to the range that contains 10 row
		Range range = Ranges.range(ss.getSelectedSheet(), selection.getRow(),
				selection.getColumn(), selection.getRow() + 9,
				selection.getColumn());

		//to effect all columns
		range = range.toRowRange(); 
		//shift exist row down and copy style from above cell 
		range.insert(InsertShift.DOWN, InsertCopyOrigin.FORMAT_LEFT_ABOVE);
	}
	
	@Listen("onClick = #deleteRows")
	public void onDeleteRows(){
		Rect selection = ss.getSelection();
		Range range = Ranges.range(ss.getSelectedSheet(), selection);

		//to effect all columns
		range = range.toRowRange(); 
		//shift exist row up 
		range.delete(DeleteShift.UP);
	}	
	
	
	@Listen("onClick = #insertColumn")
	public void onInsertColumn(){
		Rect selection = ss.getSelection();
		Range range = Ranges.range(ss.getSelectedSheet(),selection.getRow(),selection.getColumn());

		//to effect all rows
		range = range.toColumnRange(); 
		//shift exist row right and copy style from left cell 
		range.insert(InsertShift.RIGHT, InsertCopyOrigin.FORMAT_LEFT_ABOVE);
		
	}
	
	@Listen("onClick = #insertColumn3")
	public void onInsertColumn3(){
		Rect selection = ss.getSelection();
		//mark to insert to the range that contains 3 column
		Range range = Ranges.range(ss.getSelectedSheet(), selection.getRow(),
				selection.getColumn(), selection.getRow() ,
				selection.getColumn()+2);

		//to effect all rows
		range = range.toColumnRange(); 
		//shift exist row right and copy style from left cell 
		range.insert(InsertShift.RIGHT, InsertCopyOrigin.FORMAT_LEFT_ABOVE);
	}
	
	@Listen("onClick = #deleteColumns")
	public void onDeleteColumns(){
		Rect selection = ss.getSelection();
		//mark to insert to the range that contains 10 row
		Range range = Ranges.range(ss.getSelectedSheet(), selection);

		//to effect all rows
		range = range.toColumnRange(); 
		//shift exist row left 
		range.delete(DeleteShift.LEFT);
	}
	
	@Listen("onClick = #insertCellDown")
	public void onInsertCellDown(){
		Rect selection = ss.getSelection();

		Range range = Ranges.range(ss.getSelectedSheet(), selection);
		
		//shift exist row down and copy style from above cell 
		range.insert(InsertShift.DOWN, InsertCopyOrigin.FORMAT_LEFT_ABOVE);
	}
	
	@Listen("onClick = #insertCellRight")
	public void onInsertCellRight(){
		Rect selection = ss.getSelection();

		Range range = Ranges.range(ss.getSelectedSheet(), selection);
		
		//shift exist row right and copy style from left cell 
		range.insert(InsertShift.RIGHT, InsertCopyOrigin.FORMAT_LEFT_ABOVE);
	}
	
	@Listen("onClick = #deleteCellUp")
	public void onDeleteCellUp(){
		Rect selection = ss.getSelection();

		Range range = Ranges.range(ss.getSelectedSheet(), selection);
		
		//shift exist row up 
		range.delete(DeleteShift.UP);
	}
	
	@Listen("onClick = #deleteCellLeft")
	public void onDeleteCellLeft(){
		Rect selection = ss.getSelection();

		Range range = Ranges.range(ss.getSelectedSheet(), selection);
		
		//shift exist left
		range.delete(DeleteShift.LEFT);
	}
}



