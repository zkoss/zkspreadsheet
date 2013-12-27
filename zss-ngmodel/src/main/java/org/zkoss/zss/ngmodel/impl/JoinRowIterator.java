package org.zkoss.zss.ngmodel.impl;

import java.util.Iterator;
import java.util.NoSuchElementException;

import org.zkoss.zss.ngmodel.NDataRow;
import org.zkoss.zss.ngmodel.NRow;
import org.zkoss.zss.ngmodel.NSheet;

public class JoinRowIterator implements Iterator<NRow> {

	private final NSheet sheet;
	private final Iterator<NRow> iterator1;
	private final Iterator<NDataRow> iterator2;
	
	private NRow nextRow;
	private NDataRow nextDataRow;
	
	public JoinRowIterator(NSheet sheet,Iterator<NRow> iterator1,
			Iterator<NDataRow> iterator2) {
		this.sheet = sheet;
		this.iterator1 = iterator1;
		this.iterator2 = iterator2;
		
		if(iterator1.hasNext()){
			nextRow = iterator1.next();
		}
		if(iterator2.hasNext()){
			nextDataRow = iterator2.next();
		}
	}

	@Override
	public boolean hasNext() {
		return nextRow!=null || nextDataRow!=null;
	}

	@Override
	public NRow next() {
		NRow next = null;
		
		if(nextRow!=null && nextDataRow!=null){
			if(nextRow.getIndex()<nextDataRow.getIndex()){
				next = nextRow;
				nextRow = iterator1.hasNext()?iterator1.next():null;
				
			}else if(nextRow.getIndex()==nextDataRow.getIndex()){
				next = nextRow;
				nextRow = iterator1.hasNext()?iterator1.next():null;
				nextDataRow = iterator2.hasNext()?iterator2.next():null;
			}else{
				next = sheet.getRow(nextDataRow.getIndex());//return the proxy;
				nextDataRow = iterator2.hasNext()?iterator2.next():null;
			}
		}else if(nextRow!=null){
			next = nextRow;
			nextRow = iterator1.hasNext()?iterator1.next():null;
		}else if(nextDataRow!=null){
			next = sheet.getRow(nextDataRow.getIndex());//return the proxy;
			nextDataRow = iterator2.hasNext()?iterator2.next():null;
		}
		
		if(next==null){
			throw new NoSuchElementException("can't find next row");
		}
		
		return next;
	}

	@Override
	public void remove() {
		throw new UnsupportedOperationException("read only");
	}

}
