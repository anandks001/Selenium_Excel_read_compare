package util;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

public class PropertyReader {

	Properties properties;

	InputStream inputStream = null;

	public PropertyReader() {
		properties = new Properties();
		loadProperties();
	}

	private void loadProperties() {
		try {
			inputStream = new FileInputStream("src/test/java/resources/config.properties");
			properties.load(inputStream);
		} catch (IOException e) {
			System.out.print(e);
			e.printStackTrace();
		} catch (Exception e) {
			System.out.print("Unable to load config.properties" + inputStream);
			e.printStackTrace();
		}
	}

	public String readProperty(String key) {
		String value = properties.getProperty(key);
		if (key != null) {
			return value;
		} else {
			throw new RuntimeException("Unable to find property from config.properties : " + key);
		}
	}
}
