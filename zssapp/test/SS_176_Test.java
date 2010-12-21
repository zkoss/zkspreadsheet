

public class SS_176_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	setTimeout("30000");
    	selectCells(1,5,5,8);
    	click("jq('$fileMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('@menu[label=\"Export\"] a.z-menu-cnt-img')");
    	waitResponse();
    	click("jq('$exportToPdf a.z-menu-item-cnt')");
    	waitResponse();
    	
    	// TODO verify if open file window is opened
    	verifyTrue(widget(jq("@window[mode=\"highlighted\"][title=\"Export to PDF\"] div.z-window-highlighted-header")).exists());
    	
    	click("jq('$currSelection label.z-radio-cnt')");
    	waitResponse();
    	click("jq('$portrait label.z-radio-cnt')");
    	waitResponse();
    	click("jq('$export td.z-button-cm')");
    	waitResponse();
    }
}
