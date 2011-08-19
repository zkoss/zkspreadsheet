import org.zkoss.ztl.JQuery;


public class SS_015_Test extends SSAbstractTestCase {

	/**
	 * Clear Style
	 */
    @Override
    protected void executeTest() {
        JQuery cell_F_5 = getSpecifiedCell(5, 5);
        //String orgTxt = getCellContent(cell_F_5); //fail to equal, encoding seems different
        String orgTxt = "Beginning of Year";
        String orgTextColor = getCellTextColor(5, 5);

        clickCell(cell_F_5);
        clickCell(cell_F_5);
    	click("jq('$editMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('$clearStyle a.z-menu-item-cnt')");
    	waitResponse();

    	String newText = getCellText(cell_F_5);
    	waitResponse();
    	String newTextColor = getCellTextColor(5, 5);
    	
    	/**
    	 * Expected:
    	 * same cell value, different cell style
    	 */
    	verifyTrue(newText.equals(orgTxt));
    	verifyFalse(orgTextColor.equals(newTextColor));
    }
}
