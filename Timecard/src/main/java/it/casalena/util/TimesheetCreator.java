package it.casalena.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Calendar;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import it.casalena.util.Constants;

/**
 * @author iluva
 *
 */
public class TimesheetCreator {

	private static final Logger logger = LoggerFactory.getLogger(TimesheetCreator.class);

	/**
	 * Metodo statico che crea la Timesheet Anansi
	 * 
	 * @param exportGOP
	 *            ReportGop
	 * @param template
	 *            Template della Timesheet Anansi
	 * @param annoString
	 *            anno del report
	 * @param meseString
	 *            mese del report
	 */
	public static void createTimeSheets(File exportGOP, File template, String annoString, String meseString) {

		String nomeRisorsa = GOPReader.getNomeRisorsa(exportGOP);

		String basePath = FileUtils.checkEndings(Config.getProperty("directoryRootPathString"))
				+ FileUtils.checkEndings(Config.getProperty("timecardPathString"));
		checkPath(basePath, annoString, meseString);

		// Rendicontazione Mese Anno Anansi Team Nome Cognome
		File timesheet = new File(FileUtils.checkEndings(basePath) + FileUtils.checkEndings(annoString)
				+ FileUtils.checkEndings(meseString) + "Rendicontazione " + meseString + " " + annoString
				+ " Anansi Team " + nomeRisorsa + ".xlsx");
		if (timesheet.exists()) {
			logger.debug("Il file " + exportGOP.getName() + " ha già la sua relativa timesheet" + timesheet.getName());
			return;
		}

		Map<Calendar, Integer> giorni = GOPReader.extractDate(exportGOP);
		XSSFWorkbook myWorkBook = null;
		FileInputStream fsIP = null;
		FileOutputStream output_file = null;
		try {

			fsIP = new FileInputStream(template);
			myWorkBook = new XSSFWorkbook(fsIP);
			XSSFSheet sheet = myWorkBook.getSheetAt(0);
			String descrizione = sheet.getRow(Constants.rigaTimesheetDescrizione)
					.getCell(Constants.colonnaTimesheetDescrizione).getStringCellValue();
			descrizione += " " + meseString;
			sheet.getRow(Constants.rigaTimesheetDescrizione).getCell(Constants.colonnaTimesheetDescrizione)
					.setCellValue(descrizione);
			String matricola = GOPReader.getMatricolaRisorsa(exportGOP).trim();
			if ("87001187".equals(matricola) || "87001177".equals(matricola)) {
				sheet.getRow(Constants.rigaLivelloSenior).getCell(Constants.colonnaLivelloSenior).setCellValue("X");
			} else {
				sheet.getRow(Constants.rigaLivelloJunior).getCell(Constants.colonnaLivelloJunior).setCellValue("X");
			}
			for (Entry<Calendar, Integer> entry : giorni.entrySet()) {
				sheet.getRow(Constants.rigaTimesheetOreLavorate).getCell(entry.getKey().get(Calendar.DAY_OF_MONTH))
						.setCellValue(entry.getValue());
			}
			myWorkBook.setSheetName(myWorkBook.getSheetIndex(sheet), "Risorsa " + nomeRisorsa);
			XSSFFormulaEvaluator.evaluateAllFormulaCells(myWorkBook);
			output_file = new FileOutputStream(timesheet);
			myWorkBook.write(output_file);
			logger.info("Creo il timesheet " + template.getName() + " sulla base del file " + exportGOP.getName());
		} catch (Exception e) {
			logger.error("Errore nella creazione del timesheet", e);
		} finally {
			if (fsIP != null) {
				try {
					fsIP.close();
				} catch (Exception e) {
					logger.error("Errore nella chiusura del FileInputStream", e);
				}
			}
			if (myWorkBook != null) {
				try {
					myWorkBook.close();
				} catch (Exception e) {
					logger.error("Errore nella chiusura del workbook", e);
				}
			}
			if (output_file != null) {
				try {
					output_file.close();
				} catch (Exception e) {
					logger.error("Errore nella chiusura del FileOutputStream", e);
				}
			}
		}
	}

	private static boolean checkPath(String baseString, String annoString, String meseString) {

		boolean baseExists = false;

		File baseDir = new File(FileUtils.checkEndings(baseString));
		if (!baseDir.exists() || !baseDir.isDirectory()) {
			logger.error("La directory " + baseString + " non esiste! ");
			return baseExists;
		} else {
			baseExists = true;
		}
		File annoDir = new File(FileUtils.checkEndings(baseString) + FileUtils.checkEndings(annoString));
		if (!annoDir.exists()) {
			logger.info("La directory " + annoDir.getAbsolutePath() + " non esisteva. La creo.");
			annoDir.mkdir();
		} else {
			if (!annoDir.isDirectory()) {
				logger.warn("Esiste un file con lo stesso nome della directory " + annoDir.getAbsolutePath()
						+ ". Rimuovo il file e creo la directory al suo posto");
				annoDir.delete();
				annoDir.mkdir();
			}
		}
		File meseDir = new File(FileUtils.checkEndings(baseString) + FileUtils.checkEndings(annoString)
				+ FileUtils.checkEndings(meseString));
		if (!meseDir.exists()) {
			logger.info("La directory " + meseDir.getAbsolutePath() + " non esisteva. La creo.");
			meseDir.mkdir();
		} else {
			if (!meseDir.isDirectory()) {
				logger.warn("Esiste un file con lo stesso nome della directory " + meseDir.getAbsolutePath()
						+ ". Rimuovo il file e creo la directory al suo posto");
				meseDir.delete();
				meseDir.mkdir();
			}
		}

		return baseExists;
	}

}
