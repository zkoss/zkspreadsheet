package org.zkoss.zss.ngmodel.impl;

import java.util.Iterator;
import java.util.NoSuchElementException;

import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NDataCell;
import org.zkoss.zss.ngmodel.NSheet;

public class JoinCellIterator implements Iterator<NCell> {

	private final NSheet sheet;
	private final int rowIndex;
	private final Iterator<NCell> iterator1;
	private final Iterator<NDataCell> iterator2;
	
	private NCell nextCell;
	private NDataCell nextDataCell;
	
	public JoinCellIterator(NSheet sheet,int rowIndex,Iterator<NCell> iterator1,
			Iterator<NDataCell> iterator2) {
		this.sheet = sheet;
		this.rowIndex = rowIndex;
		this.iterator1 = iterator1;
		this.iterator2 = iterator2;
		
		if(iterator1.hasNext()){
			nextCell = iterator1.next();
		}
		if(iterator2.hasNext()){
			nextDataCell = iterator2.next();
		}
		checkRowIndex();
	}

	void checkRowIndex(){
		if(nextCell!=null && nextCell.getRowIndex()!=rowIndex){
			throw new IllegalStateException("row index of cell is not correct "+rowIndex+"!="+nextCell.getRowIndex());
		}
		if(nextDataCell!=null && nextDataCell.getRowIndex()!=rowIndex){
			throw new IllegalStateException("row index of data cell is not correct "+rowIndex+"!="+nextCell.getRowIndex());
		}
	}
	
	@Override
	public boolean hasNext() {
		return nextCell!=null || nextDataCell!=null;
	}

	@Override
	public NCell next() {
		NCell next = null;
		
		if(nextCell!=null && nextDataCell!=null){
			if(nextCell.getColumnIndex()<nextDataCell.getColumnIndex()){
				next = nextCell;
				nextCell = iterator1.hasNext()?iterator1.next():null;
				
			}else if(nextCell.getColumnIndex()==nextDataCell.getColumnIndex()){
				next = nextCell;
				nextCell = iterator1.hasNext()?iterator1.next():null;
				nextDataCell = iterator2.hasNext()?iterator2.next():null;
			}else{
				next = sheet.getCell(rowIndex,nextDataCell.getColumnIndex());//return the proxy;
				nextDataCell = iterator2.hasNext()?iterator2.next():null;
			}
		}else if(nextCell!=null){
			next = nextCell;
			nextCell = iterator1.hasNext()?iterator1.next():null;
		}else if(nextDataCell!=null){
			next = sheet.getCell(rowIndex,nextDataCell.getColumnIndex());//return the proxy;
			nextDataCell = iterator2.hasNext()?iterator2.next():null;
		}
		
		if(next==null){
			throw new NoSuchElementException("can't find next cell at row "+rowIndex);
		}
		checkRowIndex();
		return next;
	}

	@Override
	public void remove() {
		throw new UnsupportedOperationException("read only");
	}

}
