package org.zkoss.zss.ngmodel.impl.chart;

import java.util.Collection;
import java.util.List;

import org.zkoss.zss.ngmodel.ErrorValue;

/*package*/ class ChartDataUtil {
	
	static final int sizeOf(Object obj){
		if(obj==null){
			return 0;
		}else if(obj instanceof Collection){
			return ((Collection)obj).size();
		}else if(obj.getClass().isArray()){
			return ((Object[])obj).length;
		}else if(obj instanceof ErrorValue){
			return 0;
		}else{
			return 1;
		}
	}
	
	static final Object valueOf(Object obj,int index){
		if(obj instanceof List){//faster before collection
			return ((List)obj).get(index);
		}else if(obj instanceof Collection){
			return ((Collection)obj).toArray()[index];
		}else if(obj.getClass().isArray()){
			return ((Object[])obj)[index];
		}else if(obj instanceof ErrorValue){
			return ((ErrorValue)obj).gettErrorString();
		}else{
			return obj;
		}
	}
}
