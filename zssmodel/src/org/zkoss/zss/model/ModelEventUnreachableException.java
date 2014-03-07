package org.zkoss.zss.model;
/**
 * Indicated the ModelEvent is not reachable, which mean, book will remove the listener if it catch this exception when send event
 * @author Dennis
 *
 */
public class ModelEventUnreachableException extends RuntimeException {
	private static final long serialVersionUID = 1L;

	public ModelEventUnreachableException() {
		super();
	}

	public ModelEventUnreachableException(String message, Throwable cause) {
		super(message, cause);
	}

	public ModelEventUnreachableException(String message) {
		super(message);
	}

	public ModelEventUnreachableException(Throwable cause) {
		super(cause);
	}

	
}
