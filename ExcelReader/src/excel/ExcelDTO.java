package excel;

import java.io.Serializable;

/**
 * The Class ExcelDTO.
 * 
 * @author Sourav
 */
public class ExcelDTO implements Serializable {

	/** The Constant serialVersionUID. */
	private static final long serialVersionUID = 1L;

	/** The column name. */
	private String columnName;

	/** The column value. */
	private String columnValue;

	/**
	 * Gets the column name.
	 * 
	 * @return the column name
	 */
	public String getColumnName() {
		return columnName;
	}

	/**
	 * Sets the column name.
	 * 
	 * @param columnName
	 *            the new column name
	 */
	public void setColumnName(String columnName) {
		this.columnName = columnName;
	}

	/**
	 * Gets the column value.
	 * 
	 * @return the column value
	 */
	public String getColumnValue() {
		return columnValue;
	}

	/**
	 * Sets the column value.
	 * 
	 * @param columnValue
	 *            the new column value
	 */
	public void setColumnValue(String columnValue) {
		this.columnValue = columnValue;
	}

}
