import org.zkoss.ztl.JQuery;

public class SS_023_Test extends SSAbstractTestCase {

	@Override
	protected void executeTest() {
		JQuery cell_B_12 = getSpecifiedCell(1, 11);
		String origValue = getCellText(cell_B_12);
		JQuery cell_B_13 = getSpecifiedCell(1, 12);
		String downCellValue = getCellText(cell_B_13);
		clickCell(cell_B_12);
		clickCell(cell_B_12);
		click("$editMenu");
		waitResponse();
		mouseOver(jq("$delete"));		
		waitResponse();
		click("$deleteEntireRow");
		waitResponse();
		String newValue = getCellText(cell_B_12);
		String newDownCellValue = getCellText(cell_B_13);

		// TODO verify if pasted cell style is cleared
        verifyTrue("Original cell value=" + origValue 
        		+ ", Original down cell value=" + downCellValue, 
        		downCellValue.equals(newValue));
	}
}
