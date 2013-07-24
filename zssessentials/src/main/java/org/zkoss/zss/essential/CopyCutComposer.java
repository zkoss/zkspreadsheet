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
 * Demonstrate CellStyle, CellOperationUtil API usage
 * 
 * @author Hawk
 * 
 */
@SuppressWarnings("serial")
public class CopyCutComposer extends SelectorComposer<Component> {

	@Wire
	private Spreadsheet ss;


	@Listen("onClick = #copyButton")
	public void copy() {
		Range src = Ranges.range(ss.getSelectedSheet(), ss.getSelection());
		Range dest = Ranges.range(getResultSheet(), ss.getSelection());
		CellOperationUtil.pasteFormula(src, dest);
	}


	@Listen("onClick = #cutButton")
	public void cut() {
		Range src = Ranges.range(ss.getSelectedSheet(), ss.getSelection());
		Range dest = Ranges.range(getResultSheet(), ss.getSelection());
		CellOperationUtil.cut(src, dest);
	}

	private Sheet getResultSheet(){
		return ss.getBook().getSheetAt(1);
	}

}
