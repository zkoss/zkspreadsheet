

public class SS_169_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	verifyFalse(isWidgetVisible("$_openFileDialog"));
    	
    	click(jq("$fileMenu"));
    	waitResponse();
    	click(jq("$openFile"));
    	waitResponse();
    	
    	verifyTrue(isWidgetVisible("$_openFileDialog"));
    	
    	doubleClick(jq("$_openFileDialog $filesListbox .z-listcell").first());
    	verifyEquals(getCellText(0, 3), "Collaboration");
    }
}
