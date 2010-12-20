import org.zkoss.ztl.JQuery;
import org.zkoss.ztl.Widget;



public class SS_041_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
//    	int origChildren = widget(jq("@tab")).nChildren();
    	JQuery marketTab = jq("@tab[label=\"Market\"] span.z-tab-text");
    	int origChildren = jq("@tab").children().length();
    	click("jq('$insertMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('$insertSheet a.z-menu-item-cnt')");
    	waitResponse();
    	
    	// TODO verify if uploaed image is displayed properly
    	Widget newTab = widget(marketTab).nextSibling();
    	String content = jq(newTab).find("span").attr("textContent");
    	verifyEquals("sheet " + (origChildren + 1), content);
    	//"jq('@tab[label=\"sheet 7\"] span.z-tab-text')"
    }
}
