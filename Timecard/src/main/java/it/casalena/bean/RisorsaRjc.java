package it.casalena.bean;

import java.math.BigDecimal;

/**
 * @author iluva
 *
 */
public class RisorsaRjc {

	private String nome;
	private BigDecimal giorni;

	/**
	 * @return getter nome
	 */
	public String getNome() {
		return nome;
	}

	/**
	 * @param nome
	 *            setter
	 */
	public void setNome(String nome) {
		this.nome = nome;
	}

	/**
	 * @return getter giorni
	 */
	public BigDecimal getGiorni() {
		return giorni;
	}

	/**
	 * @param giorni
	 *            setter
	 */
	public void setGiorni(BigDecimal giorni) {
		this.giorni = giorni;
	}

}
