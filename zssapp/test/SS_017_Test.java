import org.zkoss.ztl.JQuery;

public class SS_017_Test extends SSAbstractTestCase {

	@Override
	protected void executeTest() {
		JQuery cell_B_6 = getSpecifiedCell(1, 5);
		String origValue = getCellContent(cell_B_6);
		JQuery cell_C_6 = getSpecifiedCell(2, 4);
		String rightCellValue = getCellContent(cell_C_6);
		clickCell(cell_B_6);
		clickCell(cell_B_6);
		click("jq('$editMenu button.z-menu-btn')");
		waitResponse();
		mouseOver(jq("$insert a.z-menu-cnt-img"));		
		waitResponse();
		click("jq('$shiftCellRight a.z-menu-item-cnt')");
		waitResponse();
		String newValue = getCellContent(cell_B_6);
		String newRightCellValue = getCellContent(cell_C_6);

		// TODO verify if pasted cell style is cleared
        verifyTrue("Original cell value=" + origValue 
        		+ ", Original right cell value=" + rightCellValue 
        		+ ", New right cell value=" + newRightCellValue, 
        		"".equals(newValue));
	}
}
