package it.casalena;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import ch.qos.logback.classic.LoggerContext;
import ch.qos.logback.classic.joran.JoranConfigurator;
import ch.qos.logback.core.joran.spi.JoranException;
import ch.qos.logback.core.util.StatusPrinter;
import it.casalena.util.Config;
import it.casalena.util.Constants;
import it.casalena.util.FileUtils;
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
	private static String templateFilePath = null;
	private static String logbackFileName;
	private static File template;
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
			String annoPath = FileUtils.checkEndings(GOPPathString) + FileUtils.checkEndings(annoString);
			File anno = new File(annoPath);
			if (!anno.isDirectory()) {
				continue;
			}
			logger.debug("Analizzo l'anno " + annoString);
			for (String meseString : anno.list()) {
				String mesePath = annoPath + FileUtils.checkEndings(meseString);
				File mese = new File(mesePath);
				if (!mese.isDirectory()) {
					continue;
				}
				logger.debug("Analizzo il mese " + meseString + " dell'anno " + annoString);
				for (String exportGOPString : mese.list()) {
					String exportGOPPath = mesePath + FileUtils.checkEndings(exportGOPString);
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
			}
		}
		logger.info("Operazione Completata");
	}

	private static void loadProp() {
		directoryRootPathString = FileUtils.checkEndings(Config.getProperty("directoryRootPathString"));
		templatePathString = directoryRootPathString + FileUtils.checkEndings(Config.getProperty("templatePathString"));
		timecardPathString = directoryRootPathString + FileUtils.checkEndings(Config.getProperty("timecardPathString"));
		GOPPathString = directoryRootPathString + FileUtils.checkEndings(Config.getProperty("GOPPathString"));
		templateFileName = Config.getProperty("templateFileName");
		logbackFileName = Config.getProperty("logbackFileName");
		templateFilePath = templatePathString + templateFileName;
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
		timecards = new File(timecardPathString);
		GOP = new File(GOPPathString);
		if (!template.exists() || !template.isFile()) {
			logger.error("Template file " + templateFilePath + " non esistente.");
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