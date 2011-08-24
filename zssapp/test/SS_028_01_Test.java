import org.zkoss.ztl.JQuery;


public class SS_028_01_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	//freeze some rows first
    	click("$viewMenu");
    	waitResponse();
    	mouseOver(jq("$freezeRows"));
    	waitResponse();
    	click("$freezeRow1");
    	waitResponse();
    	// TODO: Verify correct row is frozen
    	verifyTrue(jq("div.zstopblock").width() != 0);
    	JQuery p = jq("div.zstopblock");
    	int clen = p.children().length();
    	JQuery c = p.children("div:nth-child(1)");
    	int width = c.width();
    	verifyTrue("widht=" + width, jq("div.zstopblock").children("div:nth-child(1)").width() != 0);
    }
}
