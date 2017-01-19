package ch.sebpiller.jsontoxls.api;

import java.io.FileOutputStream;
import java.io.InputStream;

import org.apache.commons.io.IOUtils;
import org.junit.Before;
import org.junit.Test;

public class DefaultXlsExporterTest {
	private DefaultXlsExporter exporter;

	@Before
	public void setUp() {
		exporter = new DefaultXlsExporter();
	}

	@Test
	public void testFillTemplate() throws Exception {
		try (InputStream tpl = getClass().getResourceAsStream("/template_dummy.xlsx");
				FileOutputStream expo = new FileOutputStream("target/template_dummy_filled.xlsx")) {
			exporter.fillTemplate(tpl,
					IOUtils.toString(getClass().getResourceAsStream("/template_dummy_data1.json"), "UTF-8"), expo);
		}
	}

}
