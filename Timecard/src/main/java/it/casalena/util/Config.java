package it.casalena.util;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import it.casalena.util.Constants;

/**
 * @author iluva
 *
 */
public class Config {

	private static Properties prop = new Properties();

	private static final Logger logger = LoggerFactory.getLogger(Config.class);

	static {
		InputStream input = null;
		try {
			input = new FileInputStream(FileUtil.checkEndings(Constants.confPath) + Constants.confFileName);
			prop.load(input);
		} catch (IOException ex) {
			logger.error("Errore nel caricamento del file di properties", ex);
		} finally {
			if (input != null) {
				try {
					input.close();
				} catch (IOException e) {
					logger.error("Errore nella chiusura dell'input", e);
				}
			}
		}
	}

	/**
	 * Metodo statico che permette di ottenere il valore della property
	 * 
	 * @param key
	 *            nome della property
	 * @return valore della property
	 */
	public static String getProperty(String key) {
		return prop.getProperty(key);
	}
}
