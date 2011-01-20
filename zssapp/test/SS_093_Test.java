
public class SS_093_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	verifyFalse(jq("$_insertFormulaDialog").isVisible());
    	
        click(jq("$insertFormulaBtn"));
        waitResponse();
        
        // Verify
        verifyTrue(jq("$_insertFormulaDialog").isVisible());
    }

}
