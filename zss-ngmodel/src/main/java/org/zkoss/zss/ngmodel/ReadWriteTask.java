package org.zkoss.zss.ngmodel;

import java.util.concurrent.locks.ReadWriteLock;

/**
 * Use this task utility to help you developing you application to concurrently access book model
 * 
 * @author dennis
 *
 */
public abstract class ReadWriteTask {
	
	abstract public Object invoke();
	
	public static Object doInWriteLock(ReadWriteLock lock, ReadWriteTask task){
		lock.writeLock().lock();
		try{
			return task.invoke();
		}finally{
			lock.writeLock().unlock();
		}
	}
	
	public Object doInReadLock(ReadWriteLock lock, ReadWriteTask task){
		lock.readLock().lock();
		try{
			return task.invoke();
		}finally{
			lock.readLock().unlock();
		}
	}
}
