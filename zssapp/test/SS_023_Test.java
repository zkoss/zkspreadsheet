import org.zkoss.ztl.JQuery;

public class SS_023_Test extends SSAbstractTestCase {

	@Override
	protected void executeTest() {
		JQuery cell_B_12 = getSpecifiedCell(1, 11);
		String origValue = getCellContent(cell_B_12);
		JQuery cell_B_13 = getSpecifiedCell(1, 12);
		String downCellValue = getCellContent(cell_B_13);
		clickCell(cell_B_12);
		clickCell(cell_B_12);
		click("jq('$editMenu button.z-menu-btn')");
		waitResponse();
		mouseOver(jq("$delete a.z-menu-cnt-img"));		
		waitResponse();
		click("jq('$deleteEntireRow a.z-menu-item-cnt')");
		waitResponse();
		String newValue = getCellContent(cell_B_12);
		String newDownCellValue = getCellContent(cell_B_13);

		// TODO verify if pasted cell style is cleared
        verifyTrue("Original cell value=" + origValue 
        		+ ", Original down cell value=" + downCellValue, 
        		downCellValue.equals(newValue));
	}
}
