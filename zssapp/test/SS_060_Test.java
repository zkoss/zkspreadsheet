import org.zkoss.ztl.JQuery;


public class SS_060_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_B_8 = getSpecifiedCell(1, 7);
        clickCell(cell_B_8);
        clickCell(cell_B_8);
        click(jq("i.z-combobox-rounded-btn:eq(0)"));
        JQuery fontItem = jq("td.z-comboitem-text:eq(3)");
        
        if (fontItem != null) {
            click(fontItem);
        }
    }

}
