import org.zkoss.ztl.JQuery;

//"Select a cell
//Click Shift cells down menuitem"
public class SS_018_Test extends SSAbstractTestCase {

	@Override
	protected void executeTest() {
		JQuery cell_B_6 = getSpecifiedCell(1, 5);
		String origValue = getCellContent(cell_B_6);
		JQuery cell_B_7 = getSpecifiedCell(1, 6);
		String downCellValue = getCellContent(cell_B_7);
		clickCell(cell_B_6);
		clickCell(cell_B_6);
		click("jq('$editMenu button.z-menu-btn')");
		waitResponse();
		mouseOver(jq("$insert a.z-menu-cnt-img"));		
		waitResponse();
		click("jq('$shiftCellDown a.z-menu-item-cnt')");
		waitResponse();
		sleep(5000);
		cell_B_6 = getSpecifiedCell(1, 5);
		String newValue = getSpecifiedCell(1, 5).text();
		String newDownCellValue = getCellContent(cell_B_7);

		// TODO verify if pasted cell style is cleared
        verifyTrue("Original cell value=" + origValue 
        		+ ", Original down cell value=" + downCellValue 
        		+ ", New down cell value=" + newDownCellValue
        		+ ", new value in B6 =" + newValue, 
        		"".equals(newValue));
	}
}
