import org.zkoss.ztl.JQuery;

//select cell B6, align top, align center, align bottom
public class SS_235_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_B_6 = getSpecifiedCell(1, 5);
        clickCell(cell_B_6);
        clickCell(cell_B_6);
        
		click(jq("@dropdownbutton$valignBtn div.z-dpbutton-arrow:eq(0)"));
		waitResponse();
		click(jq("@menuitem[label=\"Top Align\"]").first());
		waitResponse();
        
		cell_B_6 = getSpecifiedCell(1, 5);
        String textAlign = cell_B_6.parent().css("vertical-align");
        
        if (textAlign != null) {
            verifyTrue("Unexcepted result: " + textAlign, textAlign.equalsIgnoreCase("top"));
        } else {
            verifyTrue("Cannot get style of specified cell!", false);
        }
        
        clickCell(cell_B_6);
        clickCell(cell_B_6);
        
		click(jq("@dropdownbutton$valignBtn div.z-dpbutton-arrow:eq(0)"));
		waitResponse();
		click(jq("@menuitem[label=\"Middle Align\"]").first());
		waitResponse();
        
		cell_B_6 = getSpecifiedCell(1, 5);
        textAlign = cell_B_6.parent().css("vertical-align");
        System.out.println("");
        
        if (textAlign != null) {
            verifyTrue("Unexcepted result: " + textAlign, textAlign.equalsIgnoreCase("middle"));
        } else {
            verifyTrue("Cannot get style of specified cell!", false);
        }

        clickCell(cell_B_6);
        clickCell(cell_B_6);
        
		click(jq("@dropdownbutton$valignBtn div.z-dpbutton-arrow:eq(0)"));
		waitResponse();
		click(jq("@menuitem[label=\"Bottom Align\"]").first());
		waitResponse();
        
		cell_B_6 = getSpecifiedCell(1, 5);
        textAlign = cell_B_6.parent().css("vertical-align");
        
        if (textAlign != null) {
            verifyTrue("Unexcepted result: " + textAlign, textAlign.equalsIgnoreCase("bottom"));
        } else {
            verifyTrue("Cannot get style of specified cell!", false);
        }
        
    }

}
