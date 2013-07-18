import org.zkoss.ztl.JQuery;

public class SS_021_Test extends SSAbstractTestCase {

	@Override
	protected void executeTest() {
		JQuery cell_F_11 = getSpecifiedCell(5, 10);
		String origValue = getCellText(cell_F_11);
		JQuery cell_G_11 = getSpecifiedCell(6, 10);
		String rightCellValue = getCellText(cell_G_11);
		JQuery cell_H_11 = getSpecifiedCell(7, 10);
		String rightRightCellValue = getCellText(cell_H_11);
		clickCell(cell_F_11);
		clickCell(cell_F_11);
		click("$editMenu");
		waitResponse();
		mouseOver(jq("$delete"));		
		waitResponse();
		click("$shiftCellLeft");
		waitResponse();
		String newValue = getCellText(cell_F_11);
		String newRightCellValue = getCellText(cell_G_11);

		// TODO verify if pasted cell style is cleared
        verifyTrue("Original cell value=" + origValue 
        		+ ", Original right cell value=" + rightCellValue 
        		+ ", Original rightRight cell value=" + rightRightCellValue 
        		+ ", New cell value=" + newValue 
        		+ ", New right cell value=" + newRightCellValue, 
        		rightCellValue.equals(newValue) && newRightCellValue.equals(rightRightCellValue));
	}
}
