package ch.sebpiller.jsontoxls.api;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.StringReader;
import java.util.ArrayList;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.commons.lang3.StringUtils;
import org.jxls.common.Context;
import org.jxls.expression.ExpressionEvaluator;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.transform.Transformer;
import org.jxls.util.JxlsHelper;
import org.jxls.util.TransformerFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonIOException;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import com.google.gson.JsonPrimitive;
import com.google.gson.JsonSyntaxException;

public class JxlsExporter implements XlsExporter {
	/** An evaluator that understand JSON from gson-api. */
	private static final class JexlSupportingJsonExpressionEvaluator extends JexlExpressionEvaluator {
		@Override
		public Object evaluate(String expression, Map<String, Object> context) {
			Object evaluate = super.evaluate(expression, context);

			if (evaluate instanceof JsonElement) {
				JsonElement json = (JsonElement) evaluate;

				if (json.isJsonObject() || json.isJsonArray()) {
					return json;
				} else if (json.isJsonPrimitive()) {
					JsonPrimitive prim = json.getAsJsonPrimitive();

					if (prim.isNumber()) {
						return prim.getAsNumber();
					} else if (prim.isBoolean()) {
						return prim.getAsBoolean();
					} else {
						return prim.getAsString();
					}
				} else if (json.isJsonNull()) {
					return null;
				}
			}

			return evaluate;
		}
	}

	/**
	 * Custom context which knows how to retrieve values from the class
	 * hierarchy of gson API.
	 */
	private static final class JsonContext extends Context {
		private final JsonElement json;

		private final Map<String, Object> resolvedValues = new LinkedHashMap<>();

		public JsonContext(JsonElement json) {
			this.json = json;
		}

		@Override
		public void putVar(String name, Object value) {
			resolvedValues.put(name, value);
		}

		@Override
		public Map<String, Object> toMap() {
			Map<String, Object> toMap = new LinkedHashMap<>();
			populateMap(toMap, json, "");
			toMap.putAll(resolvedValues);
			return toMap;
		}

		private static void populateMap(Map<String, Object> toMap, JsonElement json, String pprefix) {
			if (json.isJsonObject()) {
				JsonObject obj = json.getAsJsonObject();

				String prefix = pprefix;
				if (StringUtils.isNotBlank(prefix)) {
					prefix += ".";
				}

				for (Entry<String, JsonElement> entry : obj.entrySet()) {
					populateMap(toMap, entry.getValue(), prefix + entry.getKey());
				}
			} else if (json.isJsonArray()) {
				JsonArray array = json.getAsJsonArray();
				toMap.put(pprefix, wrapJsonArrayAsCollection(array));

				for (int i = 0; i < array.size(); i++) {
					populateMap(toMap, array.get(i), pprefix + "[" + i + "]");
				}

			} else if (json.isJsonPrimitive()) {
				toMap.put(pprefix, json.getAsString());
			} else if (json.isJsonNull()) {
				toMap.put(pprefix, null);
			}
		}

		private static Collection<JsonElement> wrapJsonArrayAsCollection(JsonArray array) {
			List<JsonElement> list = new ArrayList<>(array.size());

			for (int i = 0; i < array.size(); i++) {
				list.add(array.get(i));
			}

			return list;
		}

		@Override
		public Object getVar(String name) {
			return toMap().get(name);
		}
	}

	private static final Logger LOG = LoggerFactory.getLogger(JxlsExporter.class);

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
		Transformer transformer = TransformerFactory.createTransformer(templateStream, resultStream);
		if (transformer == null) {
			throw new IllegalStateException(
					"Cannot load XLS transformer. Please make sure a Transformer implementation is in classpath");
		}
		
		ExpressionEvaluator evaluator = new JexlSupportingJsonExpressionEvaluator();
		transformer.getTransformationConfig().setExpressionEvaluator(evaluator);
		JxlsHelper.getInstance().processTemplate(new JsonContext(json), transformer);
	}
}
