package ch.sebpiller.jsontoxls.api;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.commons.io.IOUtils;
import org.junit.Before;

public abstract class BaseXlsExporterTest {
	protected XlsExporter exporter;

	@Before
	public void setUp() {
		exporter = createExporter();
	}

	protected abstract XlsExporter createExporter();

	protected void fillTemplate(String tpl, String json, String out) throws IOException, FileNotFoundException {
		try {
			try (InputStream tplStream = getClass().getResourceAsStream(tpl);
					FileOutputStream expo = new FileOutputStream(out)) {
				exporter.fillTemplate(tplStream, IOUtils.toString(getClass().getResourceAsStream(json), "UTF-8"), expo);
			}
		} catch (Exception e) {
			throw new AssertionError("tpl: "+tpl+", json: "+json+", out: "+out, e);
		}
	}

}