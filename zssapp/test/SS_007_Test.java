import org.zkoss.ztl.JQuery;


public class SS_007_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	click("jq('$fileMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('@menu[label=\"Export\"] a.z-menu-cnt-img')");
    	waitResponse();
    	click("jq('$exportToPdf a.z-menu-item-cnt')");
    	waitResponse();
    	click("jq('$export td.z-button-cm')");
    	waitResponse();
    	
    	// TODO verify if exported pdf is generated
    }
}
