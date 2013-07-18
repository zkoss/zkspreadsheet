import org.zkoss.ztl.JQuery;

//Format>>Font>> From Calibri to Impact

public class SS_032_Test extends SSAbstractTestCase {

	/**
	 * 
	 */
    @Override
    protected void executeTest() {
        JQuery cell_B_8 = getSpecifiedCell(1, 7);
        clickCell(cell_B_8);
        clickCell(cell_B_8);

    	click(jq("$formatMenu"));
    	waitResponse();
    	mouseOver(jq("$font"));
    	waitResponse();
    	click(jq("$impactFont"));
    	waitResponse();
    	// TODO verify if cell is still selected
        String style = cell_B_8.css("font-family");
        
        if (style != null) {
            verifyTrue("Unexcepted result: " + style, "Impact".equalsIgnoreCase(style));
        } else {
            verifyTrue("Cannot get style of specified cell!", false); 
        }
   }
}
