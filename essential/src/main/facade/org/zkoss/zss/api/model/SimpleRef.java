package org.zkoss.zss.api.model;

/**
 *
 * it simple use hard reference to a instance
 * @author dennis
 *
 * @param <T>
 */
public class SimpleRef<T> implements ModelRef<T>{

	T instance;
	
	public SimpleRef(T instance){
		this.instance = instance;
	}
	
	public T get() {
		return instance;
	}
}
