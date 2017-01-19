package ch.sebpiller.jsontoxls.api;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class DefaultXlsExporter implements XlsExporter {
	private static final Logger LOG = LoggerFactory.getLogger(DefaultXlsExporter.class);

	@Override
	public void fillTemplate(InputStream templateStream, String jsonData, OutputStream resultStream) {
		try {
			doFillTemplate(templateStream, jsonData, resultStream);
		} catch (IOException e) {
			throw new ExportException("unable to export your xls sheet because of an io error: " + e, e);
		}
	}

	private void doFillTemplate(InputStream templateStream, String jsonData, OutputStream resultStream) throws IOException {
		Workbook workbook = new XSSFWorkbook(templateStream);
	}

}
