import org.zkoss.ztl.JQuery;


public class SS_034_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_J_22 = getSpecifiedCell(9, 21);
        clickCell(cell_J_22);
        clickCell(cell_J_22);
		click(jq("$formatMenu"));
		waitResponse();
		mouseOver(jq("$align"));		
		waitResponse();
		click(jq("$alignHorizontal"));
		waitResponse();
      
        cell_J_22 = getSpecifiedCell(9, 21);
        String textAlign = cell_J_22.css("text-align");
        
        if (textAlign != null) {
            verifyTrue("Unexcepted result: " + textAlign, textAlign.equalsIgnoreCase("left"));
        } else {
            verifyTrue("Cannot get style of specified cell!", false);
        }
    }

}
