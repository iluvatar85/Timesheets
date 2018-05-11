package it.casalena.util;

import java.io.File;
import java.io.FileInputStream;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import it.casalena.util.Constants;

/**
 * @author iluva
 *
 */
public class GOPReader {

	private static final Logger logger = LoggerFactory.getLogger(GOPReader.class);

	/**
	 * Metodo statico che legge in Nome ed il Cognome della risorsa dal report GOP e
	 * gli aggiunge gli appropriati spazi.
	 * 
	 * @param file
	 *            Report GOP
	 * @return Nome e Cognome
	 */
	public static String getNomeRisorsa(File file) {

		String nome = "";
		String nomeAttaccato = "";
		HSSFWorkbook myWorkBook = null;
		try {
			myWorkBook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(file)));
			HSSFSheet sheet = myWorkBook.getSheetAt(0);
			nomeAttaccato = sheet.getRow(Constants.rigaNomeRisorsa).getCell(Constants.colonnaNomeRisorsa)
					.getStringCellValue();
			if(nomeAttaccato == null || "".equals(nomeAttaccato)) {
				return nomeAttaccato;
			}
			nomeAttaccato = nomeAttaccato.replaceAll("\\s+","");
			String[] words = nomeAttaccato.split("(?=[A-Z])");
			for (int i = 0; i < words.length; i++) {
				nome += words[i];
				if (i < words.length - 1) {
					nome += " ";
				}
			}
		} catch (Exception e) {
			logger.error("Errore nella lettura del nome della risorsa", e);
		} finally {
			if (myWorkBook != null) {
				try {
					myWorkBook.close();
				} catch (Exception e) {
					logger.error("Errore nella chiusura del workbook", e);
				}
			}
		}
		return nome;
	}

	/**
	 * Metodo statico che permette di ottenere il numero di matricola della risorsa
	 * dal report GOP.
	 * 
	 * @param file
	 *            Report GOP
	 * @return Matricola
	 */
	public static String getMatricolaRisorsa(File file) {

		String matricola = "";
		HSSFWorkbook myWorkBook = null;
		try {
			myWorkBook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(file)));
			HSSFSheet sheet = myWorkBook.getSheetAt(0);
			matricola = sheet.getRow(Constants.rigaMatricolaRisorsa).getCell(Constants.colonnaMatricolaRisorsa)
					.getStringCellValue();
		} catch (Exception e) {
			logger.error("Errore nella lettura del nome della risorsa", e);
		} finally {
			if (myWorkBook != null) {
				try {
					myWorkBook.close();
				} catch (Exception e) {
					logger.error("Errore nella chiusura del workbook", e);
				}
			}
		}
		return matricola;
	}

	/**
	 * Metodo statico che consente di verificare che il report GOP sia stato
	 * posizionato nella cartella appropriata, lanciando un WARN nel caso di errato
	 * posizionamento.
	 * 
	 * TODO: automatizzare lo spostamento ed eventualmente la rielaborazione del
	 * file.
	 * 
	 * @param file
	 *            Path base delle timecard
	 * @param mese
	 *            mese della cartella
	 * @param anno
	 *            anno della cartella
	 * @return true se la data del report è consistente con la sua posizione
	 */
	public static boolean checkData(File file, String mese, String anno) {

		String meseAnno = "";
		String meseFile = "";
		String annoFile = "";
		HSSFWorkbook myWorkBook = null;
		try {
			myWorkBook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(file)));
			HSSFSheet sheet = myWorkBook.getSheetAt(0);
			meseAnno = sheet.getRow(Constants.rigaMeseAnno).getCell(Constants.colonnaMeseAnno).getStringCellValue();
			String[] stringhe = meseAnno.split(" ");
			meseFile = stringhe[0];
			annoFile = stringhe[1];
		} catch (Exception e) {
			logger.error("Errore nel controllo delle date", e);
		} finally {
			if (myWorkBook != null) {
				try {
					myWorkBook.close();
				} catch (Exception e) {
					logger.error("Errore nella chiusura del workbook", e);
				}
			}
		}

		return meseFile.toUpperCase().equals(mese.toUpperCase()) && annoFile.toUpperCase().equals(anno.toUpperCase());
	}

	/**
	 * Metodo statico che permette di estrarre tutte le date di lavoro e le ore
	 * lavorate dal report GOP
	 * 
	 * @param file
	 *            Report GOP
	 * @return mappa delle date e delle ore lavorate
	 */
	public static Map<Calendar, Integer> extractDate(File file) {
		Cell cell;
		Row row;
		Map<Calendar, Integer> giorni = new HashMap<Calendar, Integer>();
		if (file.exists()) {
			HSSFWorkbook myWorkBook = null;
			try {
				myWorkBook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(file)));
				HSSFSheet sheet = myWorkBook.getSheetAt(0);
				giorni = new HashMap<>();
				for (int i = Constants.rigaGiorni; i < sheet.getLastRowNum(); i++) {
					row = sheet.getRow(i);
					cell = row.getCell(Constants.colonnaDataLavoro);
					Date data = Constants.SDF.parse(cell.getStringCellValue());
					cell = row.getCell(Constants.colonnaOreLavoro);
					if (cell != null && !"".equals(cell.getStringCellValue())) {
						Integer ore = Integer.parseInt(cell.getStringCellValue());
						Calendar calData = new GregorianCalendar();
						calData.setTime(data);
						if (giorni.containsKey(calData)) {
							giorni.put(calData, giorni.get(calData) + ore);
						} else {
							giorni.put(calData, ore);
						}
					}
				}
			} catch (Exception e) {
				logger.error("Errore nell'estrazione delle date dal report GOP", e);
			} finally {
				if (myWorkBook != null) {
					try {
						myWorkBook.close();
					} catch (Exception e) {
						logger.error("Errore nella chiusura del workbook", e);
					}
				}
			}

		}
		return giorni;
	}
}
