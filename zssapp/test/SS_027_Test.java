import org.zkoss.ztl.JQuery;


public class SS_027_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	//freeze some rows first
    	click("jq('$viewMenu button.z-menu-btn')");
    	waitResponse();
    	mouseOver(jq("$freezeRows a.z-menu-cnt-img"));
    	waitResponse();
    	click("jq('$freezeRow5 a.z-menu-item-cnt')");
    	waitResponse();
    	verifyTrue(jq("div.zstopblock").width() != 0);
    	
    	//unfreeze all rows
    	click("jq('$viewMenu button.z-menu-btn')");
    	waitResponse();
    	mouseOver(jq("$freezeRows a.z-menu-cnt-img"));
    	waitResponse();
    	click("jq('$unfreezeRows a.z-menu-item-cnt')");
    	waitResponse();
    	
    	// TODO: Verify if all rows are unfroze
    	verifyTrue(jq("div.zstopblock").width() == 0);
    }
}
