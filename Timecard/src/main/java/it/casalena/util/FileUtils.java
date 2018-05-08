package it.casalena.util;

/**
 * @author iluva
 *
 */
public class FileUtils {

	/**
	 * Metodo statico che permette di garantire la presenza della barra finale nei
	 * percorsi dei file
	 * 
	 * @param path path da verificare
	 * @return path corretto 
	 */
	public static String checkEndings(String path) {

		if (!path.endsWith("\\")) {
			path += "\\";
		}
		return path;
	}

}
