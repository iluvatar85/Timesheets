package it.casalena;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import ch.qos.logback.classic.LoggerContext;
import ch.qos.logback.classic.joran.JoranConfigurator;
import ch.qos.logback.core.joran.spi.JoranException;
import ch.qos.logback.core.util.StatusPrinter;
import it.casalena.bean.RisorsaRjc;
import it.casalena.util.Config;
import it.casalena.util.Constants;
import it.casalena.util.FileUtil;
import it.casalena.util.GOPReader;
import it.casalena.util.TimesheetCreator;

/**
 * L'applicazione Timesheet consente di generare delle Timesheet per Anansi
 * sulla base del modello fornito e degli export delle rendicontazioni
 * dell'applicazione GOP di Autostrade.
 * 
 * I report GOP devono essere organizzati secondo una struttura
 * ANNO/MESE/file.xls
 *
 * @author iluva
 * @version 1.0.1
 * @since 2018-05-07
 */

public class Timesheets {

	private static final Logger logger = LoggerFactory.getLogger(Timesheets.class);

	private static String directoryRootPathString = null;
	private static String templatePathString = null;
	private static String timecardPathString = null;
	private static String GOPPathString = null;
	private static String templateFileName = null;
	private static String templateRagruppatoFileName = null;
	private static String templateFilePath = null;
	private static String templateRagruppatoFilePath = null;
	private static String templateRjcFilePath = null;
	private static String logbackFileName;
	private static boolean reportRjc = false;
	private static File template;
	private static File templateRagruppato;
	private static File templateRjc;
	private static File timecards;
	private static File GOP;

	/**
	 * Main method
	 * 
	 * @param args
	 *            nessun argomento necessario
	 */
	public static void main(String[] args) {
		loadProp();
		configureLogging();
		checkInitialDirs();

		for (String annoString : GOP.list()) {
			String annoPath = FileUtil.checkEndings(GOPPathString) + FileUtil.checkEndings(annoString);
			File anno = new File(annoPath);
			if (!anno.isDirectory()) {
				continue;
			}
			logger.debug("Analizzo l'anno " + annoString);
			for (String meseString : anno.list()) {
				String mesePath = annoPath + FileUtil.checkEndings(meseString);
				File mese = new File(mesePath);
				if (!mese.isDirectory()) {
					continue;
				}
				logger.debug("Analizzo il mese " + meseString + " dell'anno " + annoString);

				boolean skip = TimesheetCreator.checkTimesheetRaggruppato(annoString, meseString);
				if (skip) {
					continue;
				}
				List<RisorsaRjc> rjc = new ArrayList<RisorsaRjc>();
				// eseguo la creazione del timesheet ragruppato
				try {
					rjc = TimesheetCreator.createTimeSheetsRagruppato(mese, templateRagruppato, annoString, meseString);
				} catch (Exception e) {
					logger.error("Errore", e);
				}

				// eseguo la creazione dei timesheet singoli
				for (String exportGOPString : mese.list()) {
					String exportGOPPath = mesePath + FileUtil.checkEndings(exportGOPString);
					File exportGOP = new File(exportGOPPath);
					if (!exportGOP.isFile()) {
						continue;
					}
					logger.debug("Analizzo il file " + exportGOPString);
					if (!GOPReader.checkData(exportGOP, meseString, annoString)) {
						logger.warn("File " + exportGOPPath + " in  posizione sbagliata.");
						continue;
					}
					TimesheetCreator.createTimeSheets(exportGOP, template, annoString, meseString);
				}
				if (reportRjc) {
					logger.info("Creo il report presenze " + meseString + " " + annoString);
					TimesheetCreator.createTimeSheetsRjc(rjc, templateRjc, annoString, meseString);
				}
			}
		}
		logger.info("Operazione Completata");
		System.exit(0);
	}

	private static void loadProp() {
		directoryRootPathString = FileUtil.checkEndings(Config.getProperty("directoryRootPathString"));
		templatePathString = directoryRootPathString + FileUtil.checkEndings(Config.getProperty("templatePathString"));
		timecardPathString = directoryRootPathString + FileUtil.checkEndings(Config.getProperty("timecardPathString"));
		GOPPathString = directoryRootPathString + FileUtil.checkEndings(Config.getProperty("GOPPathString"));
		templateFileName = Config.getProperty("templateFileName");
		templateRagruppatoFileName = Config.getProperty("templateRagruppatoFileName");
		String templateRjcFileName = Config.getProperty("templateRjcFileName");
		logbackFileName = Config.getProperty("logbackFileName");
		if ("1".equals(Config.getProperty("reportRjc"))) {
			reportRjc = true;
		}
		templateFilePath = templatePathString + templateFileName;
		templateRagruppatoFilePath = templatePathString + templateRagruppatoFileName;
		templateRjcFilePath = templatePathString + templateRjcFileName;

	}

	private static void checkConfiguration() {
		LoggerContext lc = (LoggerContext) LoggerFactory.getILoggerFactory();
		StatusPrinter.print(lc);
	}

	private static void configureLogging() {

		String logbackConfigFile = Constants.confPath + logbackFileName;
		InputStream inputStream = null;
		logger.info(logbackConfigFile);
		if (logbackConfigFile != null) {
			try {
				inputStream = new FileInputStream(logbackConfigFile);
				logger.info(logbackConfigFile + "...");
				if (inputStream != null) {
					JoranConfigurator configurator = new JoranConfigurator();
					LoggerContext loggerContext = (LoggerContext) LoggerFactory.getILoggerFactory();
					loggerContext.reset();
					configurator.setContext(loggerContext);
					configurator.doConfigure(inputStream);
					checkConfiguration();
				}
			} catch (JoranException ex) {
				throw new RuntimeException(ex);
			} catch (FileNotFoundException fnfe) {
				logger.error("Errore nella riconfigurazione di logback", fnfe);
				fnfe.printStackTrace();
			}
		}

	}

	private static void checkInitialDirs() {
		template = new File(templateFilePath);
		templateRagruppato = new File(templateRagruppatoFilePath);
		templateRjc = new File(templateRjcFilePath);
		timecards = new File(timecardPathString);
		GOP = new File(GOPPathString);
		if (!template.exists() || !template.isFile()) {
			logger.error("Template file " + templateFilePath + " non esistente.");
			System.exit(1);
		}
		if (!templateRagruppato.exists() || !templateRagruppato.isFile()) {
			logger.error("Template file " + templateRagruppatoFilePath + " non esistente.");
			System.exit(1);
		}
		if (!templateRjc.exists() || !templateRjc.isFile()) {
			logger.error("Template file " + templateRjcFilePath + " non esistente.");
			System.exit(1);
		}
		if (!timecards.exists() || !timecards.isDirectory()) {
			logger.error("Timecards path " + timecardPathString + " non esistente.");
			System.exit(1);
		}
		if (!GOP.exists() || !GOP.isDirectory()) {
			logger.error("GOP path " + GOPPathString + " non esistente.");
			System.exit(1);
		}
	}
}