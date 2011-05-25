import org.zkoss.ztl.JQuery;


public class SS_171_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	setTimeout("15000");
    	click("jq('$fileMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('@menu[label=\"Export\"] a.z-menu-cnt-img')");
    	waitResponse();
    	click("jq('$exportPdf a.z-menu-item-cnt')");
    	waitResponse();
    	
    	// TODO verify if open file window is opened
    	verifyTrue(widget(jq("@window[title=\"Export to PDF\"]")).exists());
    	click("jq('$export td.z-button-cm')");
    	waitResponse();
    }
}
