package org.zkoss.zss.ngmodel.sys.dependency;

/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public interface ObjectRef extends Ref{
	
	public enum ObjectType{
		CHART, NAME, VALIDATION
	}
	public ObjectType getObjectType();
	public String getObjectId();
	public String[] getObjectIdPath();
}
