package org.zkoss.zss.essential;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.IllegalFormulaException;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.CellData;
import org.zkoss.zss.essential.util.ClientUtil;
import org.zkoss.zss.ui.Position;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Label;
import org.zkoss.zul.Textbox;

/**
 * Demonstrate cell data API usage
 * @author dennis
 *
 */
@SuppressWarnings("serial")
public class CellDataComposer extends SelectorComposer<Component> {

	@Wire
	private Label cellType;
	@Wire
	private Label cellFormatText;
	@Wire
	private Label cellEditText;
	@Wire
	private Label cellValue;
	@Wire
	private Label cellResultType;
	@Wire
	private Label cellRef;
	@Wire
	private Textbox cellEditTextBox;
	@Wire
	private Spreadsheet ss;
	
//	@Wire
//	Textbox cellFormatTextBox;
	
	
	@Listen("onCellFocus = #ss")
	public void onCellFocus(){
		Position pos = ss.getCellFocus();
		
		refreshCellInfo(pos.getRow(),pos.getColumn());		
	}
	
	private void refreshCellInfo(int row, int col){
		Range range = Ranges.range(ss.getSelectedSheet(),row,col);
		
		cellRef.setValue(Ranges.getCellReference(row, col));
		//show a cell's data
		CellData data = range.getCellData();
		cellFormatText.setValue(data.getFormatText());
		cellEditText.setValue(data.getEditText());
		cellType.setValue(data.getType().toString());
		
		Object value = data.getValue();
		cellValue.setValue(value==null?"null":(value.getClass().getSimpleName()+" : "+value));
		cellResultType.setValue(data.getResultType().toString());
		
		cellEditTextBox.setValue(data.getEditText());
		
//		cellFormatTextBox.setValue(range.getCellStyle().getDataFormat());
	}
	
	@Listen("onChange = #cellEditTextBox")
	public void onEditboxChange(){
		Position pos = ss.getCellFocus();
		Range range = Ranges.range(ss.getSelectedSheet(),pos.getRow(),pos.getColumn());
		CellData data = range.getCellData();
		if(data.validateEditText(cellEditTextBox.getValue())){
			try{
				data.setEditText(cellEditTextBox.getValue());
			}catch (IllegalFormulaException e){
				//handle illegal formula input
			}
		}else{
			ClientUtil.showWarn("not a valid value");
		}
		refreshCellInfo(pos.getRow(),pos.getColumn());
		
	}
	
	/*
	@Listen("onChange = #cellFormatTextBox")
	public void onFromatboxChange(){
		Position pos = ss.getCellFocus();
		Range range = Ranges.range(ss.getSelectedSheet(),pos.getRow(),pos.getColumn());
		range.sync(new RangeRunner() {
			public void run(Range range) {
				CellStyle oldstyle = range.getCellStyle();
				String format = cellFormatTextBox.getValue();
				if(Objects.equals(oldstyle, format)){
					return;
				}
				//don't edit old style directly, if it is shared with other cell
				EditableCellStyle newstyle = range.getCellStyleHelper().createCellStyle(oldstyle);
				newstyle.setDataFormat(format);
				range.setCellStyle(newstyle);
			}
		});
		
		refreshCellInfo(pos.getRow(),pos.getColumn());
		
	}
	*/
}



