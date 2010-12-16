import org.zkoss.ztl.JQuery;


public class SS_061_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_B_8 = getSpecifiedCell(1, 7);
        clickCell(cell_B_8);
        clickCell(cell_B_8);
        click(jq("i.z-combobox-rounded-btn:eq(1)"));
        JQuery fontSizeItem = jq(".z-combobox-rounded-pp .z-comboitem-text").eq(13);
        
        if (fontSizeItem != null) {
            click(fontSizeItem);
        }
    }

}
