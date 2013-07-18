import org.zkoss.ztl.JQuery;


public class SS_172_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	setTimeout("30000");
    	click("jq('$fileMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('@menu[label=\"Export\"] a.z-menu-cnt-img')");
    	waitResponse();
    	click("jq('$exportPdf a.z-menu-item-cnt')");
    	waitResponse();
    	
    	// TODO verify if open file window is opened
    	verifyTrue(widget(jq("@window[title=\"Export to PDF\"]")).exists());
    	click("jq('$allSheet label.z-radio-cnt')");
    	waitResponse();
    	click("jq('$export td.z-button-cm')");
    	waitResponse();
    }
}
