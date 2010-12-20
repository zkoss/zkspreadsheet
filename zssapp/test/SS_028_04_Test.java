import org.zkoss.ztl.JQuery;


public class SS_028_04_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	//freeze some rows first
    	click("jq('$viewMenu button.z-menu-btn')");
    	waitResponse();
    	mouseOver(jq("$freezeRows a.z-menu-cnt-img"));
    	waitResponse();
    	click("jq('$freezeRow4 a.z-menu-item-cnt')");
    	waitResponse();
    	// TODO: Verify correct row is frozen
    	verifyTrue(jq("div.zstopblock").width() != 0);
    	JQuery p = jq("div.zstopblock");
    	int clen = p.children().length();
    	JQuery c = p.children("div:nth-child(4)");
    	int width = c.width();
    	verifyTrue("no. of children=" + clen + ", widht=" + width, clen == 4 && jq("div.zstopblock").children("div:nth-child(4)").width() != 0);
    }
}
