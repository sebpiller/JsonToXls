package ch.sebpiller.jsontoxls.api;

import java.io.InputStream;
import java.io.OutputStream;

/**
 * Interface of objects able to take a template, a json string and produce an
 * XLS sheet with the data, using a set of predefined conventions.
 */
public interface XlsExporter {
	/**
	 * Any exception from this library gets wrapped in an
	 * {@link ExportException}.
	 */
	static final class ExportException extends RuntimeException {
		private static final long serialVersionUID = 1L;

		public ExportException() {
			super();
		}

		public ExportException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
			super(message, cause, enableSuppression, writableStackTrace);
		}

		public ExportException(String message, Throwable cause) {
			super(message, cause);
		}

		public ExportException(String message) {
			super(message);
		}

		public ExportException(Throwable cause) {
			super(cause);
		}
	}

	/**
	 * @param templateStream
	 *            A stream pointing to the template to fill. NB: don't forget to
	 *            close your own stream!
	 * @param jsonData
	 *            A (valid) JSON string containing any necessary data.
	 * @param resultStream
	 *            A stream which will contain your resulting data. NB: don't
	 *            forget to close your own stream!
	 * @throws ExportException
	 *             if anything goes terribly wrong...
	 */
	void fillTemplate(InputStream templateStream, String jsonData, OutputStream resultStream) throws ExportException;
}
