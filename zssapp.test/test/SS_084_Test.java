import org.zkoss.ztl.JQuery;

//Toolbar>>Merge cells
//merge k6:L6
public class SS_084_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_K_6 = loadCellK6();
        int cellK6Width = cell_K_6.width();
        selectCells(10, 5, 11, 5);
    	
        // Merge cells
        click(jq("$mergeCellBtn"));
        
        // Verify
        cell_K_6 = loadCellK6();
        int currentCellWidth = cell_K_6.width();
        verifyTrue("Unexcepted result: original width = " + 
                cellK6Width + ", merged cell width = " + currentCellWidth, 
                cellK6Width < currentCellWidth);
    }
    
    private JQuery loadCellK6() {
        return getSpecifiedCell(10, 5);
    }

}
