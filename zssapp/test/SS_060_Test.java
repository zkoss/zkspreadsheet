import org.zkoss.ztl.JQuery;


public class SS_060_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_B_8 = getSpecifiedCell(1, 7);
        clickCell(cell_B_8);
        clickCell(cell_B_8);
        click(jq("$fontCtrlPanel $fontFamily:visible i.z-combobox-rounded-btn"));
        JQuery fontItem = jq(".z-combobox-rounded-pp .timeFont:visible");
        
        if (fontItem != null) {
            click(fontItem);
        }
        
        cell_B_8 = getSpecifiedCell(1, 7);
        String style = cell_B_8.css("font-family");
        
        if (style != null) {
            verifyTrue("Unexcepted result: " + style, "Times New Roman".equalsIgnoreCase(style));
        } else {
            verifyTrue("Cannot get style of specified cell!", false); 
        }
    }

}
