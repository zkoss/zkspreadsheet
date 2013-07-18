import org.zkoss.ztl.JQuery;


public class SS_030_02_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	//freeze some rows first
    	click(jq("$viewMenu"));
    	waitResponse();
    	mouseOver(jq("$freezeCols"));
    	waitResponse();
    	click(jq("$freezeCol2"));
    	waitResponse();
    	// TODO: Verify correct row is frozen
    	JQuery p = jq("div.zsleftblock");
    	verifyTrue(p.width() != 0);
    	JQuery r = p.children("div:nth-child(2)");
    	int numOfCells = r.children().length();
    	int width = r.width();
    	verifyTrue("numOfCells=" + numOfCells + ", widht=" + width, r.width() != 0 && numOfCells == 2);
    }
}
