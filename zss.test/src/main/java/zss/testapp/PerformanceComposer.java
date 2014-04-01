package zss.testapp;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;

public class PerformanceComposer extends SelectorComposer<Component> {
	private static final long serialVersionUID = 1L;
	@Wire
	Spreadsheet ss;
	
	@Listen("onClick = #btn1")
	public void onClick1(){
		doActon(1);
	}
	
	@Listen("onClick = #btn40")
	public void onClick40(){
		doActon(40);
	}
	@Listen("onClick = #btn100")
	public void onClick100(){
		doActon(100);
	}
	@Listen("onClick = #btn200")
	public void onClick200(){
		doActon(200);
	}
	@Listen("onClick = #btn500")
	public void onClick500(){
		doActon(500);
	}
	
	private void doActon(int mxrow){
	    Sheet sheet = ss.getBook().getSheetAt(0);
	    long t1 = System.currentTimeMillis();
		for(int i = 0; i < mxrow; i++) {
			for(int j = 0; j < 10; j++) {
				Range range = Ranges.range(sheet, i, j);
				range.setCellValue(range.getCellData().getDoubleValue() + 5);
			}
		}
		long t2 = System.currentTimeMillis();
		System.out.println(">Row "+mxrow+" \t"+(t2-t1) +"\t(ms)");
	}
	
}
