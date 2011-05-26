

public class SS_171_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {

    	click("jq('$fileMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('@menu[label=\"Export\"] a.z-menu-cnt-img')");
    	waitResponse();
    	click("jq('$exportPdf a.z-menu-item-cnt')");
    	waitResponse();
    	
    	verifyTrue(isWidgetVisible("$_exportToPdfDialog"));
    	//TODO: after click export, how to verify "native" browser dialog (download dialog)  
    	//click("jq('$export td.z-button-cm')");
    }
}
