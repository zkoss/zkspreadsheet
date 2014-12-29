package org.zkoss.zss.app.repository.impl;

import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import org.zkoss.zss.api.model.Book;

/**
 * 
 * @author JerryChen
 *
 */

public class CollaborationInfo {
	
	class UserKey {
		private String sessionId;
		private String username;
		private Set<String> desktopIds = new HashSet<String>();
		public String getSessionId() {
			return sessionId;
		}
		public void setSessionId(String sessionId) {
			this.sessionId = sessionId;
		}
		public String getUsername() {
			return username;
		}
		public void setUsername(String username) {
			this.username = username;
		}
		public Set<String> getDesktopIds() {
			return desktopIds;
		}
		public void setDesktopIds(Set<String> desktopIds) {
			this.desktopIds = desktopIds;
		}
	}
	
	private static final String DEFAULT_USERNAME = "user";
	private static CollaborationInfo collaborationInfo;

	// book name to user set for given book name
	private Map<String, Set<String>> bookUsersRelationship = new HashMap<String, Set<String>>(10);
	// user name to user set for a book name
	private Map<String, String> userBookRelationship = new HashMap<String, String>(10);
	
	private int count = 1;
	
	private CollaborationInfo() {}
	
	public static CollaborationInfo getInstance() {
		if(collaborationInfo == null)
			collaborationInfo = new CollaborationInfo();
		return collaborationInfo;
	}

	public void setRelationship(String username, Book book) {
		if(book == null || username == null)
			return;
		
//		System.out.println(bookUsersRelationship);
//		System.out.println(userBookRelationship);
		synchronized (this) {
			if(userBookRelationship.get(username) != null)
				removeRelationship(username);
			
			String bookName = book.getBookName();
			Set<String> users = bookUsersRelationship.get(bookName);
			if (users == null) {
				users = new HashSet<String>(10);
				bookUsersRelationship.put(bookName, users);
			}
							
			users.add(username);
			userBookRelationship.put(username, bookName);
			
		}
	}

	public void removeRelationship(String username) {
		if(username == null)
			return;
		
//		System.out.println(bookUsersRelationship);
//		System.out.println(userBookRelationship);
		synchronized (this) {
			String bookName = userBookRelationship.get(username);
			
			if(bookName == null)
				return;
			
			Set<String> users = bookUsersRelationship.get(bookName);
			// remove user in users
			users.remove(username);
			userBookRelationship.put(username, null);
			if(users.isEmpty()) {
				bookUsersRelationship.remove(bookName);
				BookManagerImpl.getInstance().detachBook(new SimpleBookInfo(bookName));
			}
		}
	}
	
	public boolean isUsernameExist(String username) {
		synchronized (this) {
			return userBookRelationship.containsKey(username);
		}
	}
	
	public boolean addUsername(String username, String oldUsername) {
		synchronized (this) {
			if(userBookRelationship.containsKey(username))
				return false;
			
			if(oldUsername == username)
				return true;
			
			if(oldUsername != null)
				removeUsername(oldUsername);
			
			userBookRelationship.put(username, null);
			return true;
		}
	}
	
	public void removeUsername(String username) {
		synchronized (this) {
			if(!userBookRelationship.containsKey(username))
				return;
			
			removeRelationship(username);
			userBookRelationship.remove(username);
		}
	}
	
	public String getUsedUsernames(String bookName) {
		Set<String> names = bookUsersRelationship.get(bookName);
		return Arrays.toString(names.toArray());
	}
	
	public String getUsername(String originName) {
		synchronized (this) {
			if(originName == null) {
				while(true) {
					if(count == 1 && !isUsernameExist(DEFAULT_USERNAME)) {
						count++;
						addUsername(DEFAULT_USERNAME, null);
						return DEFAULT_USERNAME;
					} else {
						String name = DEFAULT_USERNAME + count++;
						if(!isUsernameExist(name)) {
							addUsername(name, null);
							return name;
						}
					}
				}
			} else {
				if(addUsername(originName, null)) {
					return originName;
				}
				
				int i = 2;
				while(true) {
					String name = originName + " (" + i++ + ")";
					if(addUsername(name, null))
						return name;
				}
			}
		}
		
	}
}
