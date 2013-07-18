import org.zkoss.ztl.JQuery;

public class SS_016_Test extends SSAbstractTestCase {

	/**
	 * Clear Both, style and value
	 */
	@Override
	protected void executeTest() {
		JQuery cell_F_5 = getSpecifiedCell(5, 5);
		String origValue = getCellText(5, 5);
		String orgTextColor = getCellTextColor(5, 5);
		
		clickCell(cell_F_5);
		clickCell(cell_F_5);
		click("$editMenu");
		waitResponse();
		click("$clearAll");
		waitResponse();
		String newValue = getCellText(5, 5);
		String newTextColor = getCellTextColor(5, 5);
		
		/**
		 * Expected:
		 * Set selected area style and value to empty
		 */
		verifyTrue("Original value=" + origValue + ", Cleared value=" + newValue, "".equals(newValue));
		verifyFalse(orgTextColor.equals(newTextColor));
	}
}
