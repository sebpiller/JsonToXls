package ch.sebpiller.jsontoxls.api;

import org.junit.Test;

public class JxlsExporterTest extends BaseXlsExporterTest {
	@Override
	protected XlsExporter createExporter() {
		return new JxlsExporter();
	}

	@Test
	public void testFillTemplate() throws Exception {
		fillTemplate("/template_dummy.xlsx", "/template_dummy_data1.json", "target/template_dummy_filled.xlsx");
	}

}
