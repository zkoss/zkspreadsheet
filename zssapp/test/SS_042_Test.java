import org.zkoss.ztl.JQuery;
import org.zkoss.ztl.Widget;



public class SS_042_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	click("jq('@menu[label=\"Help\"] button.z-menu-btn')");
    	waitResponse();
    	click("jq('$openCheatsheet a.z-menu-item-cnt')");
    	waitResponse();
    	
    	// TODDO check if cheatsheet window popup eixsts
    	verifyTrue("", widget(jq("$cheatsheet div.z-window-popup-cnt")).exists());
    }
}
