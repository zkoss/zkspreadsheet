

public class SS_038_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	verifyFalse(isWidgetVisible("$_insertFormulaDialog"));
    	
    	click("jq('$insertMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('$insertFormula a.z-menu-item-cnt')");
    	waitResponse();
    	
    	verifyTrue(isWidgetVisible("$_insertFormulaDialog"));
    }
}
