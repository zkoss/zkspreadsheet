import org.zkoss.ztl.JQuery;


public class SS_058_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_F_21 = getSpecifiedCell(5, 20);
        clickCell(cell_F_21);
        clickCell(cell_F_21);
        click(jq("$cutBtn"));
        
        // Verify
        verifyTrue(jq("div.zshighlight").width() != 0);
    }

}
