/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.model.util;

import java.util.concurrent.locks.ReadWriteLock;

/**
 * Use this task utility to help you developing you application to concurrently access book model
 * 
 * @author dennis
 * @since 3.5.0
 */
public abstract class ReadWriteTask {
	
	abstract public Object invoke();
	
	public Object doInWriteLock(ReadWriteLock lock){
		lock.writeLock().lock();
		try{
			return this.invoke();
		}finally{
			lock.writeLock().unlock();
		}
	}
//	
//	private static Object doInWriteLock(ReadWriteLock lock, ReadWriteTask task){
//		lock.writeLock().lock();
//		try{
//			return task.invoke();
//		}finally{
//			lock.writeLock().unlock();
//		}
//	}
	public Object doInReadLock(ReadWriteLock lock){
		lock.readLock().lock();
		try{
			return this.invoke();
		}finally{
			lock.readLock().unlock();
		}
	}
//	private static Object doInReadLock(ReadWriteLock lock, ReadWriteTask task){
//		lock.readLock().lock();
//		try{
//			return task.invoke();
//		}finally{
//			lock.readLock().unlock();
//		}
//	}
}
