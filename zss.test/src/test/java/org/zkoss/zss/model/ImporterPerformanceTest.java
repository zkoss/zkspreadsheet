package org.zkoss.zss.model;

import org.junit.*;
import org.zkoss.zss.zats.PrintStopwatch;

import java.net.URL;

public class ImporterPerformanceTest extends ImExpTestBase{

    protected URL PERFORMANCE_350K_CELL_FILE_UNDER_TEST = ImporterTest.class.getResource("book/zss-1405-350kCell.xlsx");

    @Rule
    public PrintStopwatch stopwatch = new PrintStopwatch();

    //zss-1405
    @Test
    public void performance(){
        ImExpTestUtil.loadBook(PERFORMANCE_350K_CELL_FILE_UNDER_TEST, "XLSX");
    }
}
