package org.zkoss.zss.model.sys.dependency;

/**
 * The object Reef to represent a object, and is always a dependent ref.
 * @author dennis
 * @since 3.5.0
 */
public interface ObjectRef extends Ref{
	
	public enum ObjectType{
		CHART, DATA_VALIDATION
	}
	public ObjectType getObjectType();
	public String getObjectId();
	public String[] getObjectIdPath();
}
