import org.zkoss.ztl.JQuery;
import org.zkoss.ztl.Widget;



public class SS_041_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
//    	int origChildren = widget(jq("@tab")).nChildren();
    	JQuery marketTab = jq("@tab[label=\"Market\"] span.z-tab-text");
    	int origChildren = jq("$sheets @tab").children().length();
    	click(jq("$insertMenu"));
    	waitResponse();
    	click(jq("$insertSheet"));
    	waitResponse();
    	
    	// TODO verify if uploaed image is displayed properly
    	Widget newTab = widget(marketTab).nextSibling();
    	String content = jq(newTab).find("span").text();
    	verifyEquals("sheet " + (origChildren + 1), content);
    	//"jq('@tab[label=\"sheet 7\"] span.z-tab-text')"
    }
}
