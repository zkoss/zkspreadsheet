

public class SS_038_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	verifyFalse(isWidgetVisible("$_insertFormulaDialog"));
    	
    	click(jq("$insertMenu"));
    	waitResponse();
    	click(jq("$insertFormula"));
    	waitResponse();
    	
    	verifyTrue(isWidgetVisible("$_insertFormulaDialog"));
    }
}
