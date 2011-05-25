

public class SS_169_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	verifyFalse(isWidgetVisible("$_openFileDialog"));
    	
    	click("jq('$fileMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('$openFile a.z-menu-item-cnt')");
    	waitResponse();
    	
    	verifyTrue(isWidgetVisible("$_openFileDialog"));
    	
    	doubleClick(jq("$_openFileDialog $filesListbox .z-listcell").first());
    	String text = getSpecifiedCell(0, 3).text();
    	verifyEquals(text, "Collaboration");
    }
}
