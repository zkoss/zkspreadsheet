
public class SS_047_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        click(jq("$exportToPDFBtn"));
        waitResponse();
        verifyTrue("Export to PDF dialog may be not shown.", jq("div.z-window-highlighted") != null);
    }

}
