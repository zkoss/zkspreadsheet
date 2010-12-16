import org.zkoss.ztl.JQuery;


public class SS_084_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_K_6 = loadCellK6();
        JQuery cell_L_6 = getSpecifiedCell(11, 5);
        int cellK6Width = cell_K_6.width();
        clickCell(cell_K_6);
        clickCell(cell_K_6);
        
        // Drag shift to select multiple cells.
        mouseDownAt(cell_K_6, "1,2");
        mouseMoveAt(cell_L_6, "1,2");
        mouseUpAt(cell_L_6, "1,2");
        
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
