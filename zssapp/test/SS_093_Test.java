
public class SS_093_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        click(jq("$insertFormulaBtn"));
        waitResponse();
        
        // Verify
        String titleOfPopup =  jq(".z-window-highlighted.z-window-highlighted-shadow .z-window-highlighted-header").attr("textContent");
        verifyEquals(titleOfPopup,"Insert Function");
    }

}
