

public class SS_027_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	//freeze some rows first
    	click("$viewMenu");
    	waitResponse();
    	mouseOver(jq("$freezeRows"));
    	waitResponse();
    	click("$freezeRow5");
    	waitResponse();
    	verifyTrue(jq("div.zstopblock").width() != 0);
    	
    	//unfreeze all rows
    	click("$viewMenu");
    	waitResponse();
    	mouseOver(jq("$freezeRows"));
    	waitResponse();
    	click("$unfreezeRows)");
    	waitResponse();
    	
    	// TODO: Verify if all rows are unfroze
    	verifyTrue(jq("div.zstopblock").width() == 0);
    }
}
