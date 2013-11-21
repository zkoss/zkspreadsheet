package org.zkoss.zss.ngmodel;

import junit.framework.Assert;

import org.junit.Test;
import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngapi.NRanges;
import org.zkoss.zss.ngmodel.impl.BookImpl;

public class RangeTest {

	@Test
	public void testGetRange(){
		
		NBook book = new BookImpl();
		NSheet sheet1 = book.createSheet("Sheet1");
		
		
		NRange r1 = NRanges.range(sheet1);
		Assert.assertEquals(0, r1.getRow());
		Assert.assertEquals(0, r1.getColumn());
		Assert.assertEquals(book.getMaxRowSize(), r1.getLastRow());
		Assert.assertEquals(book.getMaxColumnSize(), r1.getLastColumn());
		
		
		r1 = NRanges.range(sheet1,3,4);
		Assert.assertEquals(3, r1.getRow());
		Assert.assertEquals(4, r1.getColumn());
		Assert.assertEquals(3, r1.getLastRow());
		Assert.assertEquals(4, r1.getLastColumn());
		
		r1 = NRanges.range(sheet1,3,4,5,6);
		Assert.assertEquals(3, r1.getRow());
		Assert.assertEquals(4, r1.getColumn());
		Assert.assertEquals(5, r1.getLastRow());
		Assert.assertEquals(6, r1.getLastColumn());
	}
}
