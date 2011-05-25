import org.zkoss.ztl.JQuery;


public class SS_026_Test extends SSAbstractTestCase {

	/**
	 * Test steps:
	 * 1. Click menu View->Formula bar to hide formula bar
	 * 2. Click menu View->Formula bar to show formula bar
	 */
    @Override
    protected void executeTest() {
        JQuery cell = getSpecifiedCell(5, 20);
        clickCell(cell);
        clickCell(cell);

    	click("jq('$viewMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('$viewFormulaBar a.z-menu-item-cnt-ck')");
    	waitResponse();
    	verifyTrue(jq("$formulaBar").height() == 0);
    	click("jq('$viewMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('$viewFormulaBar a.z-menu-item-cnt-unck')");
    	waitResponse();
    	verifyTrue(jq("$formulaBar").height() != 0);
    }
}
