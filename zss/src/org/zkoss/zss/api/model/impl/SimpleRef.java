package org.zkoss.zss.api.model.impl;

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

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result
				+ ((instance == null) ? 0 : instance.hashCode());
		return result;
	}

	@Override
	@SuppressWarnings("rawtypes")
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (obj == null)
			return false;
		if (getClass() != obj.getClass())
			return false;
		SimpleRef other = (SimpleRef) obj;
		if (instance == null) {
			if (other.instance != null)
				return false;
		} else if (!instance.equals(other.instance))
			return false;
		return true;
	}
}
