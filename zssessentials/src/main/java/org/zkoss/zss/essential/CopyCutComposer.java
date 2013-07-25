package org.zkoss.zss.essential;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * Demonstrate copy & cut , CellOperationUtil API usage
 * 
 * @author Hawk
 * 
 */
@SuppressWarnings("serial")
public class CopyCutComposer extends SelectorComposer<Component> {

	@Wire
	private Spreadsheet ss;


	@Listen("onClick = #copyButton")
	public void copyByUtil() {
		Range src = Ranges.range(ss.getSelectedSheet(), ss.getSelection());
		Range dest = Ranges.range(getDestinationSheet(), ss.getSelection());
		CellOperationUtil.pasteFormula(src, dest);
	}


	@Listen("onClick = #cutButton")
	public void cutByUtil() {
		Range src = Ranges.range(ss.getSelectedSheet(), ss.getSelection());
		Range dest = Ranges.range(getDestinationSheet(), ss.getSelection());
		CellOperationUtil.cut(src, dest);
	}

	private Sheet getDestinationSheet(){
		return ss.getBook().getSheetAt(1);
	}

	//demonstrate Range API usage
	public void copy(){
		Range src = Ranges.range(ss.getSelectedSheet(), ss.getSelection());
		Range destination = Ranges.range(getDestinationSheet(), ss.getSelection());
		src.paste(destination);
	}
}
