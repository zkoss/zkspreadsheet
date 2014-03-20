package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;

import java.util.ArrayList;
import java.util.List;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;

@Ignore("manually")
public class PerformanceTest {
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		memoryInit = -1;
	}
	
	@After
	public void tearDown() throws Exception {
	}
	
	@Test
	public void testMemory() throws Exception {
		test(10,false,true);
	}
	@Test
	public void testTime() throws Exception {
		test(1,true,false);
	}
	public void test(int loop,boolean dumpTime,boolean dumpMemory) throws Exception {
		System.out.println(">>JDK:"+System.getProperty("java.runtime.version"));

		String bookPath = "book/performance.xlsx";
		
		if(dumpMemory){
			showMemoryUsage(">>>Init:"); // init
		}
		long start;
		Book book;
		Sheet sheet;
		for(int i=0;i<2;i++){
			start = System.currentTimeMillis();
			book = Util.loadBook(this,bookPath);
			sheet = book.getSheet("Sheet1");
			if(dumpTime){
				System.out.println("("+(i+1)+").Preload Book Time: " + (System.currentTimeMillis() - start));
			}
			
			book = null;
			sheet = null;
			
			if(dumpMemory){
				showMemoryUsage(">>>After Load ("+(i+1)+").book:"); // after load 1st book
			}
		}
		
		
		start = System.currentTimeMillis();
		book = Util.loadBook(this,"book/performance.xlsx");
		sheet = book.getSheet("Sheet1");
		if(dumpTime){
			System.out.println("Load Book Time: " + (System.currentTimeMillis() - start));
		}
		
		if(dumpMemory){
			showMemoryUsage(">>>After Load 2nd book:"); // after load second book
		}
		
		for(int times = 0; times < loop; times++) {
			
			start = System.currentTimeMillis();
			for(int i = 0; i < 500; i++) {
				for(int j = 0; j < 10; j++) {
					assertEquals(String.valueOf((i+1)+times*5), Ranges.range(sheet, i, j).getCellFormatText());
				}
				assertEquals(String.valueOf((i+1+times*5)*10), Ranges.range(sheet, i, 10).getCellFormatText());
				assertEquals(String.valueOf(i+1+times*5), Ranges.range(sheet, i, 11).getCellFormatText());
				assertEquals("10", Ranges.range(sheet, i, 12).getCellFormatText());
			}
			if(dumpTime){
				System.out.println("Eval Cell: " + (System.currentTimeMillis() - start));
			}
			if(dumpMemory){
				showMemoryUsage(">>>\t"+(times+1)+".eval:\t"); // after eval cells
			}
			
			start = System.currentTimeMillis();
			for(int i = 0; i < 500; i++) {
				for(int j = 0; j < 10; j++) {
					Range range = Ranges.range(sheet, i, j);
					range.setCellValue(range.getCellData().getDoubleValue() + 5);
				}
			}
			if(dumpTime){
				System.out.println("Change Value: " + (System.currentTimeMillis() - start));
			}
			if(dumpMemory){
				showMemoryUsage(">>>\t"+(times+1)+".change value:\t");
			}
			
			start = System.currentTimeMillis();
			for(int i = 0; i < 500; i++) {
				for(int j = 0; j < 10; j++) {
					assertEquals(String.valueOf((i+1+(times+1)*5)), Ranges.range(sheet, i, j).getCellFormatText());
				}
				assertEquals(String.valueOf((i+1+(times+1)*5)*10), Ranges.range(sheet, i, 10).getCellFormatText());
				assertEquals(String.valueOf(i+1+(times+1)*5), Ranges.range(sheet, i, 11).getCellFormatText());
				assertEquals("10", Ranges.range(sheet, i, 12).getCellFormatText());
			}
			if(dumpTime){
				System.out.println("Eval Cell: " + (System.currentTimeMillis() - start));
			}
			if(dumpMemory){
				showMemoryUsage(">>>\t"+(times+1)+".eval again:\t");
			}
		}
		
		if(dumpMemory){
			showMemoryUsage(">>>End:");
		}
		book = null;
		if(dumpMemory){
			showMemoryUsage(">>>Clear book:");
		}
	}
	
	
	long memoryInit = -1;
	
	private void showMemoryUsage(String msg) {
		long ms = System.currentTimeMillis();
		Runtime rt = Runtime.getRuntime();
		long last = rt.freeMemory();
		while(true) {
			rt.gc();
			Thread.yield();
			if(rt.freeMemory() == last) {
				break;
			}
			last = rt.freeMemory();
		}
		ms = System.currentTimeMillis() - ms;
		long usedHeap;
		if(memoryInit == -1) {
			memoryInit = usedHeap = rt.totalMemory() - rt.freeMemory();
		} else {
			usedHeap = rt.totalMemory() - rt.freeMemory() - memoryInit;
		}
		System.out.printf((msg==null?"":msg)+" used:  %,16d KB (GC spend %,ds)\n", usedHeap / 1024, ms / 1000L);
	}
}
