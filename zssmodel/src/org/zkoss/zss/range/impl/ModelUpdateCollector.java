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

	public void addRefs(Set<Ref> dependents) {
		ModelUpdate u = getLast();
		//optimal, merge refs, add to refs if it is just previous update
		if(u!=null){
			if(u.getType()==UpdateType.REFS){
				((Set)u.getData()).addAll(dependents);
				return;
			}else if(u.getType()==UpdateType.REF){
				 Set data = new LinkedHashSet();
				 data.add((Ref)u.getData());
				 data.addAll(dependents);
				 addModelUpdate(new ModelUpdate(UpdateType.REFS, data));
				 return;
			}
		}
		addModelUpdate(new ModelUpdate(UpdateType.REFS, new LinkedHashSet(dependents)));
	}

	public void addRef(Ref ref) {
		ModelUpdate u = getLast();
		//optimal, merge refs, add to refs if it is just previous update
		if(u!=null){
			if(u.getType()==UpdateType.REFS){
				((Set)u.getData()).add(ref);
				return;
			}else if(u.getType()==UpdateType.REF){
				 Set data = new LinkedHashSet();
				 data.add((Ref)u.getData());
				 data.add(ref);
				 addModelUpdate(new ModelUpdate(UpdateType.REFS, data));
				 return;
			}
		}
		addModelUpdate(new ModelUpdate(UpdateType.REF, ref));
	}

	public void addCellUpdate(SSheet sheet, int row, int column, int lastRow,
			int lastColumn) {
		ModelUpdate u = getLast();
		//optimal, check if it is in previous refs, ref
		if(u!=null){
			if(u.getType()==UpdateType.REFS){
				String bookName = sheet.getBook().getBookName();
				String sheetName = sheet.getSheetName();
				if(((Set)u.getData()).contains(new RefImpl(bookName,sheetName,row,column,lastRow,lastColumn))){
					//ignore if it is in previous refs
					return;
				}
			}else if(u.getType()==UpdateType.REF){
				String bookName = sheet.getBook().getBookName();
				String sheetName = sheet.getSheetName();
				if(((Ref)u.getData()).equals(new RefImpl(bookName,sheetName,row,column,lastRow,lastColumn))){
					//ignore if it is in previous refs
					return;
				}
			}else if(u.getType()==UpdateType.CELL){
				SheetRegion data = new SheetRegion(sheet,row,column,lastRow,lastColumn);
				SheetRegion prev = (SheetRegion)u.getData(); 
				if(prev.getSheet().equals(data.getSheet()) && prev.getRegion().contains(data.getRegion())){
					//ignore if it is inside previous cell
					return;
				}
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
