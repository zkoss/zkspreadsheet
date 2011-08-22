import org.zkoss.ztl.JQuery;


public class SS_015_Test extends SSAbstractTestCase {

	/**
	 * Clear Style
	 */
    @Override
    protected void executeTest() {
        JQuery cell_F_5 = getSpecifiedCell(5, 5);
        //String orgTxt = getCellText(cell_F_5); //chrome fail to equal, encoding seems different
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
    	
    	//TODO: replace "\u00A0"
    	/**
    	 * Expected:
    	 * same cell value, different cell style
    	 */
    	verifyTrue("Beginning\u00A0of\u00A0Year".equals(newText));
    	verifyFalse(orgTextColor.equals(newTextColor));
    }
}
