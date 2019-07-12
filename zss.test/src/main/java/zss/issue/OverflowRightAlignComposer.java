package zss.issue;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.*;
import org.zkoss.zss.api.*;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Vlayout;

import java.util.*;

public class OverflowRightAlignComposer extends SelectorComposer<Component> {

    @Wire
    private Spreadsheet ss;
    private LinkedList<Range> rangesToNotify = new LinkedList<Range>();

    @Override
    public void doAfterCompose(Component comp) throws Exception {
        super.doAfterCompose(comp);
        mergeAcrossFrozenColumns();
    }

    private void mergeAcrossFrozenColumns() {
        Range range = Ranges.range(ss.getSelectedSheet(),8,0,8,4);
        range.merge(false);
        range.setCellValue("A9:D9 merged by API with right alignement");
        CellOperationUtil.applyAlignment(range, CellStyle.Alignment.RIGHT);
    }

    @Listen("onClick = #fill")
    public void fill() {
        for (int i= 0 ; i < 50 ; i++) {
            getRangeWithoutRefresh(i, 0, i, 9).setCellValue(new Random().nextInt(100));
        }
        notifyRanges();
    }

    @Listen("onClick = #unfreeze")
    public void unfreeze(){
        Ranges.range(ss.getSelectedSheet()).setFreezePanel(0, 0);
    }

    @Listen("onClick = #freeze")
    public void freeze(){
        Ranges.range(ss.getSelectedSheet()).setFreezePanel(0, 3);
    }

    @Listen("onClick = #merge")
    public void merge(){
        Ranges.range(ss.getSelectedSheet(), 9, 0, 9, 3).merge(false);
        Ranges.range(ss.getSelectedSheet(), 10, 2, 10, 3).merge(false);

    }

    private void notifyRanges() {
        Ranges.range(ss.getSelectedSheet()).notifyChange();
//        while (!rangesToNotify.isEmpty()){
//            rangesToNotify.poll().notifyChange();
//        }
        rangesToNotify.clear();
    }

    private Range getRangeWithoutRefresh(int topRow, int leftColumn, int bottomRow, int rightColumn){
        Range range = Ranges.range(ss.getSelectedSheet(), topRow, leftColumn, bottomRow, rightColumn);
        range.setAutoRefresh(false);
        rangesToNotify.add(range);
        return range;
    }
}
