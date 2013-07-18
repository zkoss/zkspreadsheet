import org.zkoss.ztl.JQuery;

//"Select a cell, 
//edit -> insert -> Shift cells down"
public class SS_018_Test extends SSAbstractTestCase {

	/**
	 * Shift cell down
	 */
	@Override
	protected void executeTest() {
		JQuery cell_B_6 = getSpecifiedCell(1, 5);
		String origValue = getCellText(1, 5);

		clickCell(cell_B_6);
		clickCell(cell_B_6);
		click("$editMenu");
		waitResponse();
		mouseOver(jq("$insert"));		
		waitResponse();
		click("$shiftCellDown:visible");
		waitResponse();

		cell_B_6 = getSpecifiedCell(1, 5);
		String newValue = getCellText(1, 5);
		String newDownCellValue = getCellText(1, 6);

		/**
		 * Expect
		 * 
		 * cell text and style shift down, the original cell is empty
		 */
		verifyTrue("Shift cell: source cell shall be empty, not: " + newValue, "".equals(newValue));
        verifyTrue(newDownCellValue.equals(origValue));
	}
}
