package ch.sebpiller.jsontoxls.api;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;

@RunWith(Parameterized.class)
public class DefaultXlsExporterTechfirmTest extends BaseXlsExporterTest {	
	@Parameters
	public static Object[][] getParams() {
		List<Object[]> all = new ArrayList<>();

		new File("target/techfirm").mkdirs();
		for (File tpl : FileUtils.listFiles(new File("src/test/resources/techfirm"), new String[] { "xlsx" }, false)) {
			for (File data : FileUtils.listFiles(new File("src/test/resources/techfirm"), new String[] { "json" },
					false)) {
				all.add(new Object[] { "/techfirm/" + tpl.getName(), "/techfirm/" + data.getName(),
						"target/techfirm/" + tpl.getName() + "_" + data.getName() + ".xlsx" });
			}
		}

		return (Object[][]) all.toArray(new Object[all.size()][3]);
	}

	private String tplPath;
	private String jsonPath;
	private String outPath;

	public DefaultXlsExporterTechfirmTest(String tplPath, String jsonPath, String outPath) {
		this.tplPath = tplPath;
		this.jsonPath = jsonPath;
		this.outPath = outPath;
	}
	
	@Override
	protected XlsExporter createExporter() {
		return new JxlsExporter();
	}

	@Test
	public void test() throws Exception {
		fillTemplate(tplPath, jsonPath, outPath);
	}

}
