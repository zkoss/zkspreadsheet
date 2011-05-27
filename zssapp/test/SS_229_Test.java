import org.zkoss.ztl.JQuery;

//vertical merge g15:g17
public class SS_229_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	int g15Height = getSpecifiedCell(6, 14).height();
    	int g16Height = getSpecifiedCell(6, 15).height();
    	int g17Height = getSpecifiedCell(6, 16).height();
    	
    	selectCells(6, 14, 6, 16);
    	
        // Merge cells
        click(jq("$mergeCellBtn"));
        
        int vmergedHeight = getSpecifiedCell(6, 14).height();
        verifyTrue(g15Height+g16Height+g17Height <= vmergedHeight+2);        
    }
}
