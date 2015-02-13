package org.zkoss.zss.app.repository.impl;

import java.io.IOException;

import javax.servlet.ServletContextEvent;
import javax.servlet.ServletContextListener;

import org.zkoss.util.logging.Log;
import org.zkoss.zss.app.BookManager;
import org.zkoss.zss.app.impl.BookManagerImpl;
import org.zkoss.zss.app.repository.BookRepositoryFactory;

public class ServletContextListenerImpl implements ServletContextListener {
	
	private static final Log logger = Log.lookup(ServletContextListenerImpl.class.getName());

	@Override
	public void contextInitialized(ServletContextEvent sce) {}

	@Override
	public void contextDestroyed(ServletContextEvent sce) {
		BookManager manager = BookManagerImpl.getInstance(BookRepositoryFactory.getInstance().getRepository());
		manager.shutdownAutoFileSaving();
		try {
			manager.saveAll();
		} catch (IOException e) {
			e.printStackTrace();
			logger.error("Saving all files causes error: " + e.getMessage());
		}
	}
}
