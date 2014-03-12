package org.zkoss.zss.model;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Locale;
import java.util.Random;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.FutureTask;

import junit.framework.Assert;

import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.util.Locales;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBooks;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.impl.BookImpl;

@Ignore
public class ConcurrentTest {
	@BeforeClass
	static public void beforeClass() {
		Setup.touch();
	}
	@Before
	public void beforeTest() {
		Locales.setThreadLocal(Locale.TAIWAN);
	}
	public void testConcurrentTest(String caseName,List<FutureTask> tasks){
		int conNo = tasks.size();
		ExecutorService executor = Executors.newFixedThreadPool(conNo);
		
		for(int i=0;i<conNo;i++){
			executor.execute(tasks.get(i));
		}
		int lasti =0;
		int i = 0;
		ArrayList<Throwable> errors = new ArrayList<Throwable>();
		while(true){
			Iterator<FutureTask> iter = tasks.iterator();
			while(iter.hasNext()){
				FutureTask task = iter.next();
				try{
					if(task.isDone()){
						i++;
						iter.remove();
						task.get();//trigger if exception
					}else if(task.isCancelled()){
						i++;
						iter.remove();
						task.get();//trigger if exception
					}
				}catch(ExecutionException x){
					errors.add(x.getCause());
				}catch(InterruptedException x){
					errors.add(x);
				}
			}
			if(lasti!=i){
				lasti=i;
				System.out.println(lasti+" Tasks completed : "+caseName);
			}
			if(tasks.size()==0){
				break;
			}
			try {
				Thread.sleep(1000);
			} catch (InterruptedException e) {}
		}
		System.out.println("All Tasks completed : "+caseName);
		for(Throwable t:errors){
			t.printStackTrace();
		}
		if(errors.size()>0){
			Assert.fail(errors.get(0).toString());
		}
	}	

//	@Test
//	public void testAccessCollection(){
//		List list = new LinkedList();
//		
//		for(int i=0;i<100;i++){
//			list.add("Item"+i);
//		}
//		
//		testAccessCollection(list);
//	}
//	public void testAccessCollection(Collection col){
//		System.out.println(">>>set value done");
//		int conNo = 50;
//		ArrayList<FutureTask> tasks = new ArrayList<FutureTask>();
//		for(int i=0;i<conNo;i++){
//			FutureTask<Object> t = new FutureTask<Object>(new MyCollectionTest(col));
//			tasks.add(t);
//		}
//		testConcurrentTest("testAccessCollection",tasks);
//		
//	}
//	
//	static class MyCollectionTest implements Callable<Object>{
//		Collection list;
//		public MyCollectionTest(Collection list){
//			this.list = list;
//		}
//
//		@Override
//		public Object call() throws Exception {
//			for(Object obj:list){
//				Thread.sleep(5);
//			}
//			Iterator iter = list.iterator();
//			while(iter.hasNext()){
//				iter.next();
//				Thread.sleep(5);
//			}
//			return null;
//		}
//	}
	
	@Test
	public void testConcurrentRead(){
		SBook book = SBooks.createBook("test");
		int sheetNo = 5;
		int rowNo = 1000;
		int columnNo = 50;
		for(int s=0;s<sheetNo;s++){
			SSheet sheet = book.createSheet("Sheet"+s);
			for(int i=0;i<rowNo;i++){
				for(int j=0;j<columnNo;j++){
					sheet.getCell(i, j).setValue("("+i+","+j+")");
				}
			}
		}
		System.out.println(">>>set value done");
		int conNo = 50;
		ArrayList<FutureTask> tasks = new ArrayList<FutureTask>();
		
		for(int i=0;i<conNo;i++){
			FutureTask t = new FutureTask(new MyCurrentRead1(book,sheetNo, rowNo,columnNo));
			tasks.add(t);
		}
		testConcurrentTest("concurrent read",tasks);
	}
	
	static class MyCurrentRead1 implements Callable<Object>{
		SBook book;
		int sheetNo;
		int rowNo;
		int columnNo;
		public MyCurrentRead1(SBook book, int sheetNo, int rowNo, int columnNo){
			this.book = book;
			this.sheetNo = sheetNo;
			this.rowNo = rowNo;
			this.columnNo = columnNo;
		}
		Random r = new Random(System.currentTimeMillis());
		@Override
		public Object call() throws Exception {
			for(int s=0;s<sheetNo;s++){
				SSheet sheet = book.getSheetByName("Sheet"+s);
				for(int i=0;i<rowNo;i++){
					Object obj = sheet.getRow(i);
					for(int j=0;j<columnNo;j++){
						Assert.assertEquals("("+i+","+j+")",sheet.getCell(i, j).getValue());
					}
				}
				for(int j=0;j<columnNo;j++){
					Object obj = sheet.getColumn(j);
				}
				sheet.getCharts();
				sheet.getPictures();
			}
			return null;
		}
	}
}
