package org.zkoss.zss.essential;

import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.CellData;
import org.zkoss.zss.essential.util.ClientUtil;
import org.zkoss.zss.ui.Position;
import org.zkoss.zul.Label;
import org.zkoss.zul.Textbox;

/**
 * This class shows all the public ZK Spreadsheet you can listen to
 * @author dennis
 *
 */
public class CellDataComposer extends AbstractDemoComposer {

	@Wire
	Label cellType;
	@Wire
	Label cellFormatText;
	@Wire
	Label cellEditText;
	@Wire
	Label cellValue;
	@Wire
	Label cellResultType;
	@Wire
	Label cellRef;
	@Wire
	Textbox cellEditTextBox;
	
	
	@Listen("onCellFocused = #ss")
	public void onCellFocused(){
		Position pos = ss.getCellFocus();
		
		refreshCellInfo(pos.getRow(),pos.getColumn());		
	}
	
	private void refreshCellInfo(int row, int col){
		Range range = Ranges.range(ss.getSelectedSheet(),row,col);
		
		cellRef.setValue(Ranges.getCellReference(row, col));
		
		CellData data = range.getCellData();
		cellFormatText.setValue(data.getFormatText());
		cellEditText.setValue(data.getEditText());
		cellType.setValue(data.getType().toString());
		
		Object value = data.getValue();
		cellValue.setValue(value==null?"null":(value.getClass().getSimpleName()+" : "+value));
		cellResultType.setValue(data.getResultType().toString());
		
		cellEditTextBox.setValue(data.getEditText());
	}
	
	@Listen("onChange = #cellEditTextBox")
	public void onEditboxChange(){
		Position pos = ss.getCellFocus();
		Range range = Ranges.range(ss.getSelectedSheet(),pos.getRow(),pos.getColumn());
		CellData data = range.getCellData();
		if(data.validateEditText(cellEditTextBox.getValue())){
			data.setEditText(cellEditTextBox.getValue());
		}else{
			ClientUtil.showWarn("not a valid value");
		}
		refreshCellInfo(pos.getRow(),pos.getColumn());
		
	}
}



