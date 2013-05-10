package zss.test;

import org.zkoss.zats.mimic.ComponentAgent;
import org.zkoss.zats.mimic.operation.AuAgent;
import org.zkoss.zats.mimic.operation.AuData;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * By default, we perform all actions on currently selected sheet.
 * @author Hawk
 *
 */
public class SpreadsheetAgent {

	private ComponentAgent zss;
	private Spreadsheet spreadsheet;
	
	public SpreadsheetAgent(ComponentAgent zss) {
		this.zss = zss;
		spreadsheet = zss.as(Spreadsheet.class);
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
	
	public void copy(int topRow, int bottomRow, int leftColumn,  int rightColumn){
		zss.as(AuAgent.class)
		.post(createToolbarActionData(topRow, bottomRow,leftColumn, rightColumn).setData("act", "copy"));
	}
	
	public void paste(int rowIndex, int columnIndex){
		zss.as(AuAgent.class)
		.post(createToolbarActionData(rowIndex, rowIndex,columnIndex, columnIndex).setData("act", "paste"));
	}

	/**
	 * 
	 * @param rgbCode e.g. 009999
	 */
	public void fillBackgroundColor(int rowIndex, int columnIndex, String rgbCode){
		AuData auData = createToolbarActionData(rowIndex, rowIndex,columnIndex, columnIndex)
				.setData("act", "fillColor").setData("color", "#"+rgbCode); 
		zss.as(AuAgent.class).post(auData);
	}
	
	private AuData createToolbarActionData(int topRow, int bottomRow, int leftColumn,  int rightColumn ){
		AuData auData = new AuData("onZSSAction");
		int sheetId = spreadsheet.getBook().getSheetIndex(spreadsheet.getSelectedSheet());
		auData.setData("sheetId", Integer.toString(sheetId)).setData("tag", "toolbar")
			.setData("tRow", topRow).setData("lCol", leftColumn)
			.setData("bRow", bottomRow).setData("rCol", rightColumn);
		return auData;
	}
}
