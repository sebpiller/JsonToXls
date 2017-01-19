package ch.sebpiller.jsontoxls.api;

import java.io.FileOutputStream;

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
		exporter.fillTemplate(getClass().getResourceAsStream("/template_dummy.xlsx"),
				IOUtils.toString(getClass().getResourceAsStream("/template_dummy_data1.json"), "UTF-8"),
				new FileOutputStream("target/template_dummy_filled.xlsx"));
	}

}
