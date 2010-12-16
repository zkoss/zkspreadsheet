import org.zkoss.ztl.JQuery;

public class SS_022_Test extends SSAbstractTestCase {

	@Override
	protected void executeTest() {
		JQuery cell_F_11 = getSpecifiedCell(5, 10);
		String origValue = getCellContent(cell_F_11);
		JQuery cell_F_12 = getSpecifiedCell(5, 11);
		String downCellValue = getCellContent(cell_F_12);
		JQuery cell_F_13 = getSpecifiedCell(5, 12);
		String downDownCellValue = getCellContent(cell_F_13);
		clickCell(cell_F_11);
		clickCell(cell_F_11);
		click("jq('$editMenu button.z-menu-btn')");
		waitResponse();
		mouseOver(jq("$delete a.z-menu-cnt-img"));		
		waitResponse();
		click("jq('$shiftCellUp a.z-menu-item-cnt')");
		waitResponse();
		String newValue = getCellContent(cell_F_11);
		String newDownCellValue = getCellContent(cell_F_12);

		// TODO verify if pasted cell style is cleared
        verifyTrue("Original cell value=" + origValue 
        		+ ", Original down cell value=" + downCellValue 
        		+ ", Original downDown cell value=" + downDownCellValue 
        		+ ", New cell value=" + newValue 
        		+ ", New down cell value=" + newDownCellValue, 
        		downCellValue.equals(newValue) && newDownCellValue.equals(downDownCellValue));
	}
}
