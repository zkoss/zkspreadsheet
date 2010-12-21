import org.zkoss.ztl.JQuery;
import org.zkoss.ztl.util.ColorVerifingHelper;


public class SS_090_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_O_6 = getSpecifiedCell(14, 5);
        click(jq("$gridlinesCheckbox"));
        waitResponse();
        cell_O_6 = getSpecifiedCell(14, 5);
        verifyTrue(ColorVerifingHelper.isEqualColor("rgb(208, 215, 233)", 
                cell_O_6.parent().css("border-bottom-color")));
    }

}
