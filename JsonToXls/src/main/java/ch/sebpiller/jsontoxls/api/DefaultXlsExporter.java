package ch.sebpiller.jsontoxls.api;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.StringReader;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

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

public class DefaultXlsExporter implements XlsExporter {
	private static final Logger LOG = LoggerFactory.getLogger(DefaultXlsExporter.class);
	private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("\\$\\{(.*)\\}");

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
			Sheet sheet = workbook.getSheetAt(0);

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
					Matcher matcher = PLACEHOLDER_PATTERN.matcher(value);

					if (matcher.matches()) {
						String jsonPath = matcher.group(1);
						JsonElement jsonElement = getJsonElementAtPath(json, jsonPath);

						// TODO handle date, number, etc... (only string yet)
						if (jsonElement == null) {
							currentCell.setCellValue("ERROR!");
						} else {
							currentCell.setCellValue(jsonElement.getAsString());
						}
					}
				}
			}

			workbook.write(resultStream);
		}
	}

	private JsonElement getJsonElementAtPath(JsonElement json, String jsonPath) {
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
