

public class SS_029_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	//freeze some rows first
    	click("jq('$viewMenu button.z-menu-btn')");
    	waitResponse();
    	mouseOver(jq("$freezeCols a.z-menu-cnt-img"));
    	waitResponse();
    	click("jq('$freezeCol5 a.z-menu-item-cnt')");
    	waitResponse();
    	verifyTrue(jq("div.zsleftblock").width() != 0);
    	
    	//unfreeze all rows
    	click("jq('$viewMenu button.z-menu-btn')");
    	waitResponse();
    	mouseOver(jq("$freezeCols a.z-menu-cnt-img"));
    	waitResponse();
    	click("jq('$unfreezeCols a.z-menu-item-cnt')");
    	waitResponse();
    	
    	// TODO: Verify if all columns are unfroze
    	verifyTrue(jq("div.zsleftblock").width() == 0);
    }
}
