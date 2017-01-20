package ch.sebpiller.jsontoxls.api;

import org.junit.Test;

public class DefaultXlsExporterTest extends BaseXlsExporterTest {
	@Test
	public void testFillTemplate() throws Exception {
		fillTemplate("/template_dummy.xlsx", "/template_dummy_data1.json", "target/template_dummy_filled.xlsx");
	}

}
