package it.casalena.util;

import java.text.SimpleDateFormat;

/**
 * @author iluva
 *
 */
@SuppressWarnings("javadoc")
public class Constants {

	public static final String FORMAT = "dd/MM/yyyy";
	public static final SimpleDateFormat SDF = new SimpleDateFormat(FORMAT);
	
	public static final String confPath = "C:\\appConf\\Anansi\\";
	public static final String confFileName = "timesheet.properties";
	
	// GOP
	public static final int rigaDataInizio = 5;
	public static final int rigaGiorni = 9;
	public static final int rigaLivelloJunior = 11;
	public static final int rigaLivelloSenior = 10;
	public static final int rigaMatricolaRisorsa = 4;
	public static final int rigaMeseAnno = 0;
	public static final int rigaNomeRisorsa = 3;
	public static final int colonnaDataInizio = 1;
	public static final int colonnaDataLavoro = 0;
	public static final int colonnaLivelloJunior = 11;
	public static final int colonnaLivelloSenior = 11;
	public static final int colonnaMatricolaRisorsa = 1;
	public static final int colonnaMeseAnno = 0;
	public static final int colonnaOreLavoro = 3;
	public static final int colonnaNomeRisorsa = 1;

	// TimeSheet
	public static final int rigaTimesheetDescrizione = 3;
	public static final int rigaTimesheetNome = 8;
	public static final int rigaTimesheetOreLavorate = 13;
	public static final int colonnaTimesheetDescrizione = 7;
	public static final int colonnaTimesheetNome = 5;
	
	
	
	
}
