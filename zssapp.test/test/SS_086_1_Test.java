import org.zkoss.ztl.JQuery;


public class SS_086_1_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_K_6 = loadCellK6();
        clickCell(cell_K_6);
        clickCell(cell_K_6);
        click(jq("$insertHyperlinkBtn"));
        waitResponse();
        String urlValue = "http://ja.wikipedia.org/wiki";
        type(jq("$addrCombobox input.z-combobox-inp"), urlValue);
        waitResponse();
        
        keyDownNative(SHIFT);
        keyDownNative(TAB);
        keyUpNative(TAB);
        keyUpNative(SHIFT);
        
        type(jq("$displayHyperlink"), "Wiki Japan");
        
        click(jq("$okBtn"));
        waitResponse();
        
        cell_K_6 = loadCellK6();
        JQuery aTag = cell_K_6.find("a");
        
        if (aTag != null) {
            String href = aTag.attr("href");
            verifyEquals("Unexpected result: " + href, "http://ja.wikipedia.org/wiki", href);
        } else {
            verifyTrue("Cannot get value of specified cell!", false);
        }
    }
    
    private JQuery loadCellK6() {
        return getSpecifiedCell(10, 5);
    }

}
