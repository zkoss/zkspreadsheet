


public class SS_042_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	click(jq("@menu[label=\"Help\"]"));
    	waitResponse();
    	click(jq("$openCheatsheet"));
    	waitResponse();
    	
    	// TODDO check if cheatsheet window popup eixsts
    	verifyTrue("", widget(jq("$cheatsheet div.z-window-popup-cnt")).exists());
    }
}
