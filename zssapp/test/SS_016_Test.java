import org.zkoss.ztl.JQuery;

public class SS_016_Test extends SSAbstractTestCase {

	@Override
	protected void executeTest() {
		JQuery cell_F_4 = getSpecifiedCell(5, 4);
		String origValue = getCellContent(cell_F_4);
		clickCell(cell_F_4);
		clickCell(cell_F_4);
		click("jq('$editMenu button.z-menu-btn')");
		waitResponse();
		click("jq('$clearAll a.z-menu-item-cnt')");
		waitResponse();
		String newValue = getCellContent(cell_F_4);

		// TODO verify if pasted cell style is cleared
        verifyTrue("Original value=" + origValue + ", Cleared value=" + newValue, "".equals(newValue));
	}
}
