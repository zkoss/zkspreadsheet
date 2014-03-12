package org.zkoss.zss.model;
/**
 * Indicated the Listener is not reachable, which mean, book could remove the listener if it catch this exception when send event
 * @author Dennis
 *
 */
public class ModelEventListenerUnreachableException extends RuntimeException {
	private static final long serialVersionUID = 1L;

	public ModelEventListenerUnreachableException() {
		super();
	}

	public ModelEventListenerUnreachableException(String message, Throwable cause) {
		super(message, cause);
	}

	public ModelEventListenerUnreachableException(String message) {
		super(message);
	}

	public ModelEventListenerUnreachableException(Throwable cause) {
		super(cause);
	}

	
}
