package ch.sebpiller.jsontoxls.api;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.StringReader;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.JsonElement;
import com.google.gson.JsonIOException;
import com.google.gson.JsonParser;
import com.google.gson.JsonSyntaxException;

/**
 * A default implementation for TechFirm SA, where the conventions are defined
 * to look like the placeholders mechanism you find sometimes in xml:
 * 
 * <code>${myobject.myfield}, ${myobject.myarray[4].anotherfield}, etc.</code>
 * 
 * It supports rows replications when iterating over a collection. The notation
 * should be "${myobject.myarray[].myfield}" (no index means a repetition).
 * 
 * Actually, the repetition can be only of 1 row. If needed, a further
 * enhancement could be to define a convention to say: hey framework, copy those
 * x lines when you find an iteration.
 * 
 * @deprecated use {@link JxlsExporter} instead. It was an (unsuccessfull)
 *             attempt to implement what JXLS does out-of-the-box.
 */
@Deprecated
public class DefaultXlsExporter implements XlsExporter {
	private static final Logger LOG = LoggerFactory.getLogger(DefaultXlsExporter.class);
	private static final Pattern PLACEHOLDER_PATTERN = anyPlaceholder();

	private static Pattern anyPlaceholder() {
		return placeholder(null);
	}

	private static Pattern placeholder(String v) {
		String var = v == null ? "[a-zA-Z0-9\\.\\[\\]]*" : Pattern.quote(v);

		return Pattern.compile("\\$\\{?(" + var + ")\\}?");
	}

	@Override
	public void fillTemplate(InputStream templateStream, String jsonData, OutputStream resultStream) {
		JsonElement json;
		try {
			json = new JsonParser().parse(new StringReader(jsonData));
			LOG.debug("successfully parsed json data: {}", json);
		} catch (JsonIOException | JsonSyntaxException e) {
			throw new ExportException("could not parse the json data: " + e, e);
		}

		try {
			doFillTemplate(templateStream, json, resultStream);
			LOG.debug("exported successfully!");
		} catch (IOException e) {
			throw new ExportException("unable to export your xls sheet because of an io error: " + e, e);
		}
	}

	private void doFillTemplate(InputStream templateStream, JsonElement json, OutputStream resultStream)
			throws IOException {
		try (Workbook workbook = new XSSFWorkbook(templateStream)) {

			for (int sheet = 0; sheet < workbook.getNumberOfSheets(); sheet++) {
				processSheet(json, workbook.getSheetAt(sheet));
			}

			workbook.write(resultStream);
		}
	}

	private void processSheet(JsonElement json, Sheet sheet) {
		for (int row = sheet.getFirstRowNum(); row < sheet.getLastRowNum(); row++) {
			Row currentRow = sheet.getRow(row);

			if (currentRow == null) {
				LOG.debug("found a null row at index {}... skipping", row);
				continue;
			}

			for (int col = currentRow.getFirstCellNum(); col < currentRow.getLastCellNum(); col++) {
				Cell currentCell = currentRow.getCell(col);

				if (currentCell == null) {
					LOG.debug("found a null cell at row {}, index {}... skipping", row, col);
					continue;
				}

				String value = currentCell.getStringCellValue();
				if (StringUtils.isBlank(value)) {
					LOG.debug("found a blank cell at row {}, index {}... skipping", row, col);
					continue;
				}

				Matcher matcher = PLACEHOLDER_PATTERN.matcher(value);
				String result = value;
				System.out.println("getCellValue(" + col + "," + row + ") " + result);

				while (matcher.find()) {
					String jsonPath = matcher.group(1);
					JsonElement jsonElement = getJsonElementAtPath(json, jsonPath);

					// TODO handle date, number, etc... (only string
					// yet)
					if (jsonElement != null) {
						result = result.replaceAll(placeholder(jsonPath).toString(), jsonElement.getAsString());
					}
				}

				System.out.println("setCellValue(" + col + "," + row + ") " + result);
				currentCell.setCellValue(result);
			}
		}
	}

	private JsonElement getJsonElementAtPath(JsonElement json, String jsonPath) {
		LOG.debug("searching path {}", jsonPath);
		String[] tokens = jsonPath.split("\\.");

		JsonElement currentElement = json;
		for (String token : tokens) {
			currentElement = currentElement.getAsJsonObject().get(token);

			if (currentElement == null) {
				LOG.warn("path {} not found!", jsonPath);
				break;
			}
		}

		return currentElement;
	}
}
