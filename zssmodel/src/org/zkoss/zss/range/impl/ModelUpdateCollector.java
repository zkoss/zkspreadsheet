package org.zkoss.zss.range.impl;

import java.util.Collections;
import java.util.LinkedHashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Set;

import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.SheetRegion;
import org.zkoss.zss.model.impl.RefImpl;
import org.zkoss.zss.model.sys.dependency.Ref;
import org.zkoss.zss.range.impl.ModelUpdate.UpdateType;

public class ModelUpdateCollector {
	static ThreadLocal<ModelUpdateCollector> current = new ThreadLocal<ModelUpdateCollector>();

	private List<ModelUpdate> updates;

	public ModelUpdateCollector() {
	}

	public static ModelUpdateCollector setCurrent(ModelUpdateCollector ctx) {
		ModelUpdateCollector old = current.get();
		current.set(ctx);
		return old;
	}

	public static ModelUpdateCollector getCurrent() {
		return current.get();
	}

	public void addModelUpdate(ModelUpdate mu) {
		if (updates == null) {
			updates = new LinkedList<ModelUpdate>();
		}
		updates.add(mu);
	}

	public List<ModelUpdate> getModelUpdates() {
		
		if(updates==null){
			return Collections.EMPTY_LIST;
		}
		return Collections.unmodifiableList(updates);
	}
	
	private ModelUpdate getLast(){
		return (updates==null||updates.size()==0)?null:updates.get(updates.size()-1);
	}
	private void removeLast(){
		if(updates!=null && updates.size()>0){
			updates.remove(updates.size()-1);
		}
	}

	public void addRefs(Set<Ref> dependents) {
		ModelUpdate last = getLast();
		//optimal, merge refs, add to refs if it is just previous update
		if(last!=null){
			if(last.getType()==UpdateType.REFS){
				((Set)last.getData()).addAll(dependents);
				return;
			}else if(last.getType()==UpdateType.REF){
				 Set data = new LinkedHashSet();
				 data.add((Ref)last.getData());
				 data.addAll(dependents);
				 
				 removeLast();
				 
				 addModelUpdate(new ModelUpdate(UpdateType.REFS, data));
				 return;
			}
		}
		addModelUpdate(new ModelUpdate(UpdateType.REFS, new LinkedHashSet(dependents)));
	}

	public void addRef(Ref ref) {
		ModelUpdate last = getLast();
		//optimal, merge refs, add to refs if it is just previous update
		if(last!=null){
			if(last.getType()==UpdateType.REFS){
				((Set)last.getData()).add(ref);
				return;
			}else if(last.getType()==UpdateType.REF){
				 Set<Ref> data = new LinkedHashSet<Ref>();
				 data.add((Ref)last.getData());
				 data.add(ref);
				 
				 removeLast();
				 
				 addModelUpdate(new ModelUpdate(UpdateType.REFS, data));
				 return;
			}
		}
		addModelUpdate(new ModelUpdate(UpdateType.REF, ref));
	}

	public void addCellUpdate(SSheet sheet, int row, int column, int lastRow,
			int lastColumn) {
		ModelUpdate last = getLast();
		//optimal, check if it is in previous refs, ref
		if(last!=null){
			if(last.getType()==UpdateType.REFS){
				String bookName = sheet.getBook().getBookName();
				String sheetName = sheet.getSheetName();
				if(((Set)last.getData()).contains(new RefImpl(bookName,sheetName,row,column,lastRow,lastColumn))){
					//ignore if it is in previous refs
					return;
				}
			}else if(last.getType()==UpdateType.REF){
				String bookName = sheet.getBook().getBookName();
				String sheetName = sheet.getSheetName();
				if(((Ref)last.getData()).equals(new RefImpl(bookName,sheetName,row,column,lastRow,lastColumn))){
					//ignore if it is in previous refs
					return;
				}
			}else if(last.getType()==UpdateType.CELLS){
				SheetRegion data = new SheetRegion(sheet,row,column,lastRow,lastColumn);
				((Set<SheetRegion>)last.getData()).add(data); 
				return;
			}else if(last.getType()==UpdateType.CELL){
				Set<SheetRegion> data = new LinkedHashSet<SheetRegion>();
				data.add((SheetRegion)last.getData());
				data.add(new SheetRegion(sheet,row,column,lastRow,lastColumn));
				 
				removeLast();
				
				addModelUpdate(new ModelUpdate(UpdateType.CELLS, data));
			}
		}
		addModelUpdate(new ModelUpdate(UpdateType.CELL, new SheetRegion(sheet,
				row, column, lastRow, lastColumn)));
	}

	public void addMergeChange(SSheet sheet, CellRegion original,
			CellRegion changeTo) {
		addModelUpdate(new ModelUpdate(UpdateType.MERGE, new MergeUpdate(sheet,
				original, changeTo)));
	}

	public void addInsertDeleteUpdate(SSheet sheet, boolean inserted,
			boolean isRow, int index, int lastIndex) {
		addModelUpdate(new ModelUpdate(
				UpdateType.INSERT_DELETE,
				new InsertDeleteUpdate(sheet, inserted, isRow, index, lastIndex)));
	}

}
