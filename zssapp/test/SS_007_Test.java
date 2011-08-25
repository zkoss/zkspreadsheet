

public class SS_007_Test extends SSAbstractTestCase {

	/**
	 * Export PDF
	 */
    @Override
    protected void executeTest() {
    	click(jq("$fileMenu"));
    	mouseOver("@menu[label=\"Export\"]");
    	waitResponse();
    	click("$exportPdf");
    	waitResponse();
   
    	//TODO: verify if exported pdf is generated
    	//click("jq('$export td.z-button-cm')");
    	//waitResponse();
    }
}
