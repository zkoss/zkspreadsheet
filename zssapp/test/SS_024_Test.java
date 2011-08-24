import org.zkoss.ztl.JQuery;

public class SS_024_Test extends SSAbstractTestCase {

	@Override
	protected void executeTest() {
		JQuery cell_F_11 = getSpecifiedCell(5, 10);
		String origValue = getCellText(cell_F_11);
		JQuery cell_G_11 = getSpecifiedCell(6, 10);
		String rightCellValue = getCellText(cell_G_11);
		clickCell(cell_F_11);
		clickCell(cell_F_11);
		click("$editMenu");
		waitResponse();
		mouseOver(jq("$delete"));		
		waitResponse();
		click("$deleteEntireColumn");
		waitResponse();
		String newValue = getCellText(cell_F_11);

		// TODO verify if pasted cell style is cleared
        verifyTrue("Original cell value=" + origValue 
        		+ ", Original right cell value=" + rightCellValue, 
        		rightCellValue.equals(newValue));
	}
}
