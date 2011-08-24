

public class SS_029_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	//freeze some rows first
    	click(jq("$viewMenu"));
    	waitResponse();
    	mouseOver(jq("$freezeCols"));
    	waitResponse();
    	click(jq("$freezeCol5"));
    	waitResponse();
    	verifyTrue(jq("div.zsleftblock").width() != 0);
    	
    	//unfreeze all rows
    	click(jq("$viewMenu"));
    	waitResponse();
    	mouseOver(jq("$freezeCols"));
    	waitResponse();
    	click(jq("$unfreezeCols"));
    	waitResponse();
    	
    	// TODO: Verify if all columns are unfroze
    	verifyTrue(jq("div.zsleftblock").width() == 0);
    }
}
