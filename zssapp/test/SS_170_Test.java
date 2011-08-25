

public class SS_170_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	selectCells(1,5,5,8);
    	click(jq("$fileMenu"));
    	waitResponse();
    	click(jq("@menu[label=\"Export\"]"));
    	waitResponse();
    	click(jq("$exportPdf"));
    	waitResponse();
    	
    	// TODO verify if open file window is opened
    	verifyTrue(widget(jq("@window[title=\"Export to PDF\"]")).exists());
    	//manually review
//    	click("jq('$currSelection label.z-radio-cnt')");
//    	waitResponse();
//    	click("jq('$export td.z-button-cm')");
//    	waitResponse();
    }
}
