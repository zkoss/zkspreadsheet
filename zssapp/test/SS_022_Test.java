import org.zkoss.ztl.JQuery;

//"Select a cell
//Click Shift cells up menuitem"
public class SS_022_Test extends SSAbstractTestCase {

	
	/**
	 * Shift cell up
	 */
	@Override
	protected void executeTest() {
		JQuery cell_F_11 = getSpecifiedCell(5, 10);
		String origValue = getCellText(cell_F_11);
		JQuery cell_F_12 = getSpecifiedCell(5, 11);
		String downCellValue = getCellText(cell_F_12);
		JQuery cell_F_13 = getSpecifiedCell(5, 12);
		String downDownCellValue = getCellText(cell_F_13);
		clickCell(cell_F_11);
		clickCell(cell_F_11);
		click("jq('$editMenu button.z-menu-btn')");
		waitResponse();
		mouseOver(jq("$delete a.z-menu-cnt-img"));		
		waitResponse();
		click("jq('$shiftCellUp a.z-menu-item-cnt')");
		waitResponse();
		String newValue = getCellText(cell_F_11);
		String newDownCellValue = getCellText(cell_F_12);

        verifyTrue("Original cell value=" + origValue 
        		+ "\n, Original down cell value=" + downCellValue 
        		+ "\n, Original downDown cell value=" + downDownCellValue 
        		+ "\n, New cell value=" + newValue 
        		+ "\n, New down cell value=" + newDownCellValue, 
        		downCellValue.equals(newValue) && newDownCellValue.equals(downDownCellValue));
	}
}
