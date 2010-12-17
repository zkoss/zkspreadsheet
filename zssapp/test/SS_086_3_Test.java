import org.zkoss.ztl.JQuery;


public class SS_086_3_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_K_6 = loadCellK6();
        clickCell(cell_K_6);
        clickCell(cell_K_6);
        click(jq("$insertHyperlinkBtn"));
        waitResponse();
        click(jq("$docBtn"));
        waitResponse();
        click(jq("$refSheet .z-treecell:eq(0)"));
        waitResponse();
        click(jq("$okBtn"));
        waitResponse();
        
        cell_K_6 = loadCellK6();
        JQuery aTag = cell_K_6.find("a");
        
        if (aTag != null) {
            String href = aTag.attr("href");
            verifyEquals("Unexpected result: " + href, "Input!A1", href);
        } else {
            verifyTrue("Cannot get value of specified cell!", false);
        }
    }
    
    private JQuery loadCellK6() {
        return getSpecifiedCell(10, 5);
    }

}
