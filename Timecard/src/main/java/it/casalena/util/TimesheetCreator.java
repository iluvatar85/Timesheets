package it.casalena.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import it.casalena.bean.RisorsaRjc;

/**
 * @author iluva
 *
 */
public class TimesheetCreator {

	private static final Logger logger = LoggerFactory.getLogger(TimesheetCreator.class);

	@SuppressWarnings("javadoc")
	public static void createTimeSheetsRjc(List<RisorsaRjc> rjc, File templateRjc, String annoString,
			String meseString) {
		// TODO Auto-generated method stub
		if (rjc == null) {
			return;
		}
		String basePath = FileUtil.checkEndings(Config.getProperty("directoryRootPathString"))
				+ FileUtil.checkEndings(Config.getProperty("timecardPathString"));
		checkPath(basePath, annoString, meseString);

		// Rendicontazione Mese Anno Anansi Team Nome Cognome
		File presenze = new File(FileUtil.checkEndings(basePath) + FileUtil.checkEndings(annoString)
				+ FileUtil.checkEndings(meseString) + "Presenze Colonica " + meseString + " " + annoString + ".xlsx");

		XSSFWorkbook myWorkBook = null;
		FileInputStream fsIP = null;
		FileOutputStream output_file = null;
		try {

			if (!presenze.exists()) {
				// copiare il file template nella posizione desiderata per poi modificarlo
				FileUtils.copyFile(templateRjc, presenze);
			}

			fsIP = new FileInputStream(presenze);
			myWorkBook = new XSSFWorkbook(fsIP);
			XSSFSheet sheet = myWorkBook.getSheetAt(0);
			int rowNum = Constants.rigaPresenzePartenza;
			sheet.getRow(Constants.rigaPresenseMese).getCell(Constants.colonnaPresenzeMese)
					.setCellValue(meseString.toUpperCase());
			for (RisorsaRjc risorsa : rjc) {
				XSSFRow row = sheet.getRow(rowNum++);
				row.getCell(Constants.colonnaPresenzeCognome).setCellValue(risorsa.getNome());
				XSSFCell giorniCell = row.getCell(Constants.colonnaPresenzeGiorni);
				giorniCell.setCellValue(risorsa.getGiorni().floatValue());
			}
			XSSFFormulaEvaluator.evaluateAllFormulaCells(myWorkBook);
			output_file = new FileOutputStream(presenze);
			myWorkBook.write(output_file);
			logger.info("Creo il file presenze " + presenze.getName());
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

	@SuppressWarnings("javadoc")
	public static List<RisorsaRjc> createTimeSheetsRagruppato(File mese, File templateRagruppato, String annoString,
			String meseString) throws IOException {

		String basePath = FileUtil.checkEndings(Config.getProperty("directoryRootPathString"))
				+ FileUtil.checkEndings(Config.getProperty("timecardPathString"));
		checkPath(basePath, annoString, meseString);

		// Rendicontazione Mese Anno Anansi Team Nome Cognome
		File timesheet = new File(
				FileUtil.checkEndings(basePath) + FileUtil.checkEndings(annoString) + FileUtil.checkEndings(meseString)
						+ "Rendicontazione " + meseString + " " + annoString + " Anansi Team.xlsx");
		// File timecardRagruppate = new
		// File(FileUtil.checkEndings(mese.getAbsolutePath()) + "Rendicontazione "
		// + meseString + " " + annoString + " Anansi Team.xlsx");
		if (!timesheet.exists()) {
			// copiare il file template nella posizione desiderata per poi modificarlo
			FileUtils.copyFile(templateRagruppato, timesheet);

		}
		List<RisorsaRjc> risorse = new ArrayList<RisorsaRjc>();
		int i = 0;
		XSSFWorkbook myWorkBook = null;
		FileInputStream fsIP = null;
		FileOutputStream output_file = null;
		try {
			fsIP = new FileInputStream(timesheet);
			myWorkBook = new XSSFWorkbook(fsIP);
			for (String exportGOPString : mese.list()) {
				String exportGOPPath = FileUtil.checkEndings(mese.getAbsolutePath()) + exportGOPString;
				File exportGOP = new File(exportGOPPath);
				if (!exportGOP.isFile()) {
					continue;
				}
				logger.debug("Analizzo il file " + exportGOPString);
				if (!GOPReader.checkData(exportGOP, meseString, annoString)) {
					logger.warn("File " + exportGOPPath + " in  posizione sbagliata.");
					continue;
				}

				String nomeRisorsa = GOPReader.getNomeRisorsa(exportGOP);

				Map<Calendar, Integer> giorni = GOPReader.extractDate(exportGOP);

				XSSFSheet sheet = myWorkBook.getSheetAt(i);
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

				BigDecimal oreTotali = new BigDecimal(0);

				for (Entry<Calendar, Integer> entry : giorni.entrySet()) {
					sheet.getRow(Constants.rigaTimesheetOreLavorate).getCell(entry.getKey().get(Calendar.DAY_OF_MONTH))
							.setCellValue(entry.getValue());
					oreTotali = oreTotali.add(new BigDecimal(entry.getValue()));
				}
				myWorkBook.setSheetName(myWorkBook.getSheetIndex(sheet), "Risorsa " + nomeRisorsa);
				XSSFFormulaEvaluator.evaluateAllFormulaCells(myWorkBook);

				if (!"87001187".equals(matricola)) {
					RisorsaRjc risorsa = new RisorsaRjc();
					risorsa.setNome(nomeRisorsa);
					BigDecimal giorniRjc = new BigDecimal(0);
					if (oreTotali.doubleValue() > 0) {
						giorniRjc = oreTotali.divide(new BigDecimal(8), 2, BigDecimal.ROUND_HALF_UP);
					}
					risorsa.setGiorni(giorniRjc);
					risorse.add(risorsa);
				}

				logger.info(
						"Modifico il timesheet " + timesheet.getName() + " sulla base del file " + exportGOP.getName());

				i++;
			}

			for (int j = i; j < myWorkBook.getNumberOfSheets(); i++) {
				myWorkBook.removeSheetAt(j);
			}
			output_file = new FileOutputStream(timesheet);
			myWorkBook.write(output_file);

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
		return risorse;
	}

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

		String basePath = FileUtil.checkEndings(Config.getProperty("directoryRootPathString"))
				+ FileUtil.checkEndings(Config.getProperty("timecardPathString"));
		checkPath(basePath, annoString, meseString);

		// Rendicontazione Mese Anno Anansi Team Nome Cognome
		File timesheet = new File(
				FileUtil.checkEndings(basePath) + FileUtil.checkEndings(annoString) + FileUtil.checkEndings(meseString)
						+ "Rendicontazione " + meseString + " " + annoString + " Anansi Team " + nomeRisorsa + ".xlsx");
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
			logger.info("Creo il timesheet " + timesheet.getName() + " sulla base del file " + exportGOP.getName());
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

		File baseDir = new File(FileUtil.checkEndings(baseString));
		if (!baseDir.exists() || !baseDir.isDirectory()) {
			logger.error("La directory " + baseString + " non esiste! ");
			return baseExists;
		} else {
			baseExists = true;
		}
		File annoDir = new File(FileUtil.checkEndings(baseString) + FileUtil.checkEndings(annoString));
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
		File meseDir = new File(FileUtil.checkEndings(baseString) + FileUtil.checkEndings(annoString)
				+ FileUtil.checkEndings(meseString));
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

	/**
	 * @param annoString
	 *            anno analizzato
	 * @param meseString
	 *            mese analizzato
	 * @return booleano che indica se saltare il mese in analisi
	 */
	public static boolean checkTimesheetRaggruppato(String annoString, String meseString) {
		boolean skip = false;

		String basePath = FileUtil.checkEndings(Config.getProperty("directoryRootPathString"))
				+ FileUtil.checkEndings(Config.getProperty("timecardPathString"));
		String gopMesePath = FileUtil.checkEndings(Config.getProperty("directoryRootPathString"))
				+ FileUtil.checkEndings(Config.getProperty("GOPPathString")) + FileUtil.checkEndings(annoString)
				+ FileUtil.checkEndings(meseString);
		String timesheetMesePath = FileUtil.checkEndings(basePath) + FileUtil.checkEndings(annoString)
				+ FileUtil.checkEndings(meseString);

		File gopMese = new File(gopMesePath);
		File timesheetGlobale = new File(
				timesheetMesePath + "Rendicontazione " + meseString + " " + annoString + " Anansi Team.xlsx");
		if (timesheetGlobale.exists()) {
			FileInputStream tsgfis = null;
			XSSFWorkbook timesheetRagruppati = null;
			try {
				tsgfis = new FileInputStream(timesheetGlobale);
				timesheetRagruppati = new XSSFWorkbook(tsgfis);
				int risorseGiàRaggruppate = timesheetRagruppati.getNumberOfSheets();
				int gopPresenti = gopMese.list().length;
				if (risorseGiàRaggruppate >= gopPresenti) {
					logger.debug("Tutte le timesheet sono state già riunite per il mese di " + meseString + " "
							+ annoString);
					skip = true;
				} else {
					// eliminare il file delle timesheet raggruppate
					FileUtils.deleteQuietly(timesheetGlobale);
				}
			} catch (Exception e) {
				logger.error("Errore", e);
			} finally {
				try {
					tsgfis.close();
				} catch (Exception e) {
					logger.error("Errore", e);
				}
				try {
					timesheetRagruppati.close();
				} catch (Exception e) {
					logger.error("Errore", e);
				}
			}
		}
		return skip;
	}

}
