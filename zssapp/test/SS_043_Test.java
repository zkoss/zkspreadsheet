import org.zkoss.ztl.JQuery;
import org.zkoss.ztl.Widget;



public class SS_043_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	click("jq('@menu[label=\"Help\"] button.z-menu-btn')");
    	waitResponse();
    	click("jq('$forum a.z-menu-item-cnt')");
    	waitResponse();
    	
    	// TODO How to verify? Can't access other window or browser tab
    }
}
