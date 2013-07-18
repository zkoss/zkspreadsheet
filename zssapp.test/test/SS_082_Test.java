import org.zkoss.ztl.JQuery;


public class SS_082_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_B_8 = getSpecifiedCell(1, 7);
        clickCell(cell_B_8);
        clickCell(cell_B_8);
        
		click(jq("@dropdownbutton$halignBtn div.z-dpbutton-arrow:eq(0)"));
		waitResponse();
		click(jq("@menuitem[label=\"Center Text\"]").first());
		waitResponse();
                
        cell_B_8 = getSpecifiedCell(1, 7);
        String textAlign = cell_B_8.css("text-align");
        
        if (textAlign != null) {
            verifyTrue("Unexcepted result: " + textAlign, textAlign.equalsIgnoreCase("center"));
        } else {
            verifyTrue("Cannot get style of specified cell!", false);
        }
    }

}
