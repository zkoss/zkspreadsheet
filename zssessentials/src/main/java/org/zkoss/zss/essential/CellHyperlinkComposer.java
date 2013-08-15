package org.zkoss.zss.essential;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.IllegalFormulaException;
import org.zkoss.zss.api.CellRef;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.CellData;
import org.zkoss.zss.api.model.Hyperlink;
import org.zkoss.zss.essential.util.ClientUtil;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Label;
import org.zkoss.zul.Textbox;

/**
 * Demonstrate cell data API usage
 * @author dennis
 *
 */
@SuppressWarnings("serial")
public class CellHyperlinkComposer extends SelectorComposer<Component> {

	@Wire
	private Label cellFormatText;
	@Wire
	private Label hyperlinkType;
	@Wire
	private Label hyperlinkLabel;
	@Wire
	private Label hyperlinkAddress;
	@Wire
	private Label cellRef;
	@Wire
	private Spreadsheet ss;
	
//	@Wire
//	Textbox cellFormatTextBox;
	
	
	@Listen("onCellFocus = #ss")
	public void onCellFocus(){
		CellRef pos = ss.getCellFocus();
		
		refreshCellInfo(pos.getRow(),pos.getColumn());		
	}
	
	private void refreshCellInfo(int row, int col){
		Range range = Ranges.range(ss.getSelectedSheet(),row,col);
		
		cellRef.setValue(Ranges.getCellReferenceString(row, col));
		//show a cell's data
		CellData data = range.getCellData();
		cellFormatText.setValue(data.getFormatText());
		
		Hyperlink link = range.getCellHyperlink();
		if(link!=null){
			hyperlinkType.setValue(link.getType().toString());
			hyperlinkLabel.setValue(link.getLabel());
			hyperlinkAddress.setValue(link.getAddress());
		}else{
			hyperlinkType.setValue("");
			hyperlinkLabel.setValue("");
			hyperlinkAddress.setValue("");
		}
	}
}



