package org.zkoss.zss.test.selenium.testcases.function;

import org.junit.Test;
import org.openqa.selenium.Keys;
import org.zkoss.zss.test.selenium.*;
import org.zkoss.zss.test.selenium.entity.*;
import org.zkoss.zss.test.selenium.util.MyPngCropper;

import static org.zkoss.zss.test.selenium.util.MyPngCropper.TOP_BOTTOM_CROPPER;


public class MergeTest extends ZSSTestCase {

    private static final String HIDDEN = "/merge/1390-merge-hidden.zul";

    /*
     * verify the importing result
     */
    @Test
    public void unhideRow() throws Exception {
        getTo(HIDDEN);
        SpreadsheetWidget ss = focusSheet();
        SheetCtrlWidget ctrl = ss.getSheetCtrl();
        click(ctrl.getCell("A1"));
        System.out.println(ctrl.getCell("A2").attr("width"));
    }

    @Test
    public void unhideColumn() throws Exception {
        getTo(HIDDEN);
        SpreadsheetWidget ss = focusSheet();
        SheetCtrlWidget ctrl = ss.getSheetCtrl();
        click(ctrl.getCell("G5"));
    }

}





