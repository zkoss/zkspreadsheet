package zss.test;

import org.zkoss.zats.mimic.ComponentAgent;
import org.zkoss.zats.mimic.operation.AuAgent;
import org.zkoss.zats.mimic.operation.AuData;
import org.zkoss.zss.ui.Spreadsheet;

public class SpreadsheetAgent {

	private ComponentAgent zss;
	
	public SpreadsheetAgent(ComponentAgent zss) {
		this.zss = zss;
	}
	
	public void selectSheet(int index){
		AuData auData = new AuData("onZSSSelectSheet");
		auData.setData("sheetId", Integer.toString(index)).setData("cache", false)
			.setData("row", -1).setData("col", -1).setData("left", -1).setData("right", -1)
			.setData("top", -1).setData("bottom", -1)
			.setData("hleft", -1).setData("hright", -1)
			.setData("htop", -1).setData("hbottom", -1)
			.setData("frow", -1).setData("fcol", -1);
		zss.as(AuAgent.class).post(auData);
	}
	
	public void selectSheet(String name){
		int sheetIndex = zss.as(Spreadsheet.class).getBook().getSheetIndex(name);
		if (sheetIndex >=0){
			selectSheet(sheetIndex);
		}
	}
	
	public void copy(int rowIndex, int columnIndex){
		AuData auData = new AuData("onZSSAction");
		auData.setData("sheetId", "0").setData("tag", "toolbar").setData("act", "copy")
			.setData("tRow", rowIndex).setData("lCol", columnIndex)
			.setData("bRow", rowIndex).setData("rCol", columnIndex);
		zss.as(AuAgent.class).post(auData);
	}
}
