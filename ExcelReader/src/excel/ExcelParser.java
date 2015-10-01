package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

/**
 * The Class ExcelParser.
 * 
 * @author Sourav
 */
public class ExcelParser {

	/** The Constant LOG. */
	private static final Logger LOG = Logger.getLogger(ExcelParser.class);

	/** The map. */
	private static Map<Integer, List<ExcelDTO>> map;

	/** The Constant EMPTY_STRING. */
	private static final String EMPTY_STRING = "";

	/**
	 * Read.
	 * 
	 * @param excelFile
	 *            the excel file
	 * @param lastCellNum
	 *            the last cell number(signifies number of columns to read from the excel)
	 * @param sheetNumber
	 *            the sheet number(Starts from 1)
	 * @return the map
	 * @throws IOException
	 *             Signals that an I/O exception has occurred.
	 * @throws SAXException
	 *             the SAX exception
	 * @throws OpenXML4JException
	 *             the OpenXML4J exception
	 */
	public static Map<Integer, List<ExcelDTO>> read(File excelFile,
			int lastCellNum, int sheetNumber) throws IOException, SAXException,
			OpenXML4JException {

		LOG.info("Beginning method [read]...");

		FileInputStream inputStream = null;
		Workbook workBook = null;
		ArrayList<String> headerList;
		ArrayList<ExcelDTO> listExcelDTO;
		DateFormat df = new SimpleDateFormat("yyyy-MM-dd");
		map = new HashMap<Integer, List<ExcelDTO>>();

		if (excelFile.getName().endsWith("xlsx")) {
			map = readXLSX(excelFile, lastCellNum, sheetNumber);
			return map;
		}

		try {
			inputStream = new FileInputStream(excelFile);
			workBook = WorkbookFactory.create(inputStream);
			Sheet sheet = workBook.getSheetAt(sheetNumber - 1);
			removeEmptyRowsFromExcelSheet(sheet);

			// Loop over column and lines
			headerList = new ArrayList<String>();

			for (int rowIndex = 0; rowIndex < sheet.getPhysicalNumberOfRows(); rowIndex++) {
				listExcelDTO = new ArrayList<ExcelDTO>();
				Row row = sheet.getRow(rowIndex);

				for (int columnIndex = 0; columnIndex < lastCellNum; columnIndex++) {
					ExcelDTO objExcelDTO = new ExcelDTO();
					Cell cell = row.getCell(columnIndex);

					if (null == cell) {
						objExcelDTO.setColumnValue("");
					} else {
						if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
							objExcelDTO.setColumnValue("");
						} else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
							objExcelDTO.setColumnValue(cell
									.getStringCellValue());
						} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
							if (DateUtil.isCellDateFormatted(cell)) {
								String celldata = df.format(cell
										.getDateCellValue());
								objExcelDTO.setColumnValue(celldata);
							} else {
								DecimalFormat decimalformat = new DecimalFormat(
										"0.000");
								String data = decimalformat.format(cell
										.getNumericCellValue());
								objExcelDTO.setColumnValue(data);
							}
						} else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {

							switch (cell.getCachedFormulaResultType()) {
							case Cell.CELL_TYPE_BLANK:
								objExcelDTO.setColumnValue("");
								break;
							case Cell.CELL_TYPE_NUMERIC:
								if (DateUtil.isCellDateFormatted(cell)) {
									String celldata = df.format(cell
											.getDateCellValue());
									objExcelDTO.setColumnValue(celldata);
								} else {
									DecimalFormat decimalformat = new DecimalFormat(
											"0.000");
									String data = decimalformat.format(cell
											.getNumericCellValue());
									objExcelDTO.setColumnValue(data);
								}
								break;

							case Cell.CELL_TYPE_STRING:
								objExcelDTO.setColumnValue(cell
										.getStringCellValue());
								break;
							}
						}

					}

					if (rowIndex == 0) {
						// headers
						objExcelDTO.setColumnName("HEADER");
						headerList.add(cell.getStringCellValue());
					} else {
						objExcelDTO.setColumnName(headerList.get(columnIndex));
					}

					listExcelDTO.add(objExcelDTO);

				}

				map.put(rowIndex, listExcelDTO);

			}
			LOG.info("Completed method [read].");

		} catch (Exception ex) {
			LOG.error("Error in method [read]:::" + ex);
		} finally {
			if (null != inputStream) {
				inputStream.close();
			}
		}

		return map;
	}

	/**
	 * Read xlsx.
	 * 
	 * @param excelFile
	 *            the excel file
	 * @param lastCellNum
	 *            the last cell number
	 * @param sheetNumber
	 *            the sheet number
	 * @return the map
	 * @throws IOException
	 *             Signals that an I/O exception has occurred.
	 * @throws OpenXML4JException
	 *             the OpenXML4J exception
	 * @throws SAXException
	 *             the SAX exception
	 */
	private static Map<Integer, List<ExcelDTO>> readXLSX(File excelFile,
			int lastCellNum, int sheetNumber) throws IOException,
			OpenXML4JException, SAXException {

		LOG.info("Beginning method [readXLSX]...");
		try {
			FileInputStream inputStream = null;
			OPCPackage opcPackage = null;
			inputStream = new FileInputStream(excelFile);
			opcPackage = OPCPackage.open(inputStream);

			XSSFReader r = new XSSFReader(opcPackage);
			StylesTable styles = r.getStylesTable();
			ReadOnlySharedStringsTable sharedStrings = new ReadOnlySharedStringsTable(
					opcPackage);
			map = new HashMap<Integer, List<ExcelDTO>>();
			ContentHandler sheetContentsHandler = new SheetHandler(
					sharedStrings, lastCellNum, styles, map);

			XMLReader parser = XMLReaderFactory.createXMLReader();

			parser.setContentHandler(sheetContentsHandler);
			InputStream sheet2 = r.getSheet(findSheetId(sheetNumber));
			InputSource sheetSource = new InputSource(sheet2);
			parser.parse(sheetSource);
			sheet2.close();
			map = completeMap(map);

			LOG.info("Completed [readXLSX].");

		} catch (Exception ex) {
			LOG.error("Error in method [readXLSX]:::" + ex);
		}

		return map;

	}

	/**
	 * Find sheet id.
	 * 
	 * @param sheetNumber
	 *            the sheet number
	 * @return the string
	 */
	private static String findSheetId(int sheetNumber) {
		String relId = "rId";
		return (relId + sheetNumber);
	}

	/**
	 * Complete map.
	 * 
	 * @param map
	 *            the map
	 * @return the map
	 */
	private static Map<Integer, List<ExcelDTO>> completeMap(
			Map<Integer, List<ExcelDTO>> map) {

		Map<Integer, List<ExcelDTO>> newExcelMap = new HashMap<Integer, List<ExcelDTO>>();
		Set<Integer> keySet = map.keySet();
		ArrayList<Integer> list = new ArrayList<Integer>(keySet);
		Collections.sort(list); //

		Iterator<Integer> itExcelRow = list.iterator();
		Integer headerKey = itExcelRow.next();
		List<ExcelDTO> excelHeaderRow = map.get(headerKey);
		List<ExcelDTO> excelRow = null;

		// set the header
		newExcelMap.put(headerKey, excelHeaderRow);

		while (itExcelRow.hasNext()) {

			Integer keyVal = itExcelRow.next();
			excelRow = map.get(keyVal);

			// get excel row value
			for (ExcelDTO excelHeaderColumnValue : excelHeaderRow) {
				if (!isColumnValuePresent(excelRow, excelHeaderColumnValue)) {
					ExcelDTO newExcelDTO = new ExcelDTO();
					newExcelDTO.setColumnName(excelHeaderColumnValue
							.getColumnValue());
					newExcelDTO.setColumnValue("");
					excelRow.add(newExcelDTO);
				}
			}

			// skip the row insert if all columns of the row is empty
			if (!isEmptyRow(excelRow)) {
				newExcelMap.put(keyVal, excelRow);
			}
		}
		return newExcelMap;
	}

	/**
	 * Checks if the column value is present.
	 * 
	 * @param excelRow
	 *            the excel row
	 * @param excelHeaderColumnValue
	 *            the excel header column value
	 * @return true, if the column value is present
	 */
	private static boolean isColumnValuePresent(List<ExcelDTO> excelRow,
			ExcelDTO excelHeaderColumnValue) {

		for (ExcelDTO dto : excelRow) {
			if (null != excelHeaderColumnValue
					&& excelHeaderColumnValue.getColumnValue()
							.equalsIgnoreCase(dto.getColumnName())) {
				return true;
			}
		}
		return false;

	}

	/**
	 * Checks if the row is empty.
	 * 
	 * @param excelRow
	 *            the excel row
	 * @return true, if the row is empty
	 */
	private static boolean isEmptyRow(List<ExcelDTO> excelRow) {
		boolean isEmptyRow = true;
		for (ExcelDTO excelColumn : excelRow) {
			if (!isEmptyColumn(excelColumn)) {
				isEmptyRow = false;
				break;
			}
		}
		return isEmptyRow;
	}

	/**
	 * Checks if the column is empty.
	 * 
	 * @param excelColumn
	 *            the excel column
	 * @return true, if the column is empty
	 */
	private static boolean isEmptyColumn(ExcelDTO excelColumn) {
		return (null == excelColumn.getColumnValue() || EMPTY_STRING
				.equalsIgnoreCase(excelColumn.getColumnValue()));
	}

	/**
	 * Removes the empty rows from the excel sheet.
	 * 
	 * @param sheet
	 *            the sheet
	 */
	private static void removeEmptyRowsFromExcelSheet(Sheet sheet) {
		boolean stop = false;
		boolean nonBlankRowFound;
		Row lastRow = null;
		Cell cell = null;

		while (!stop) {
			nonBlankRowFound = false;
			lastRow = sheet.getRow(sheet.getLastRowNum());
			for (int cellIndex = lastRow.getFirstCellNum(); cellIndex <= lastRow
					.getLastCellNum(); cellIndex++) {
				cell = lastRow.getCell(cellIndex);
				if (cell != null
						&& lastRow.getCell(cellIndex).getCellType() != Cell.CELL_TYPE_BLANK) {
					nonBlankRowFound = true;
				}
			}
			if (nonBlankRowFound) {
				stop = true;
			} else {
				sheet.removeRow(lastRow);
			}
		}
	}
}

class SheetHandler extends DefaultHandler {
	enum xssfDataType {
		BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER, NO_STYLE
	}

	private static final Logger LOG = Logger.getLogger(SheetHandler.class);

	private List<ExcelDTO> listExcelDTO = new ArrayList<ExcelDTO>();
	private StringBuffer value = new StringBuffer();
	private StylesTable stylesTable;

	/** * Table with unique strings */
	private ReadOnlySharedStringsTable sharedStringsTable;

	private List<String> headerList = new ArrayList<String>();

	/** * Number of columns to read starting with leftmost */
	private final int minColumnCount = 5;

	// Set when V start element is seen
	private boolean vExists;

	// Set when cell start element is seen
	// used when cell close element is seen.
	private xssfDataType nextDataType;

	// Used to format numeric cell values
	private int formatIndex;
	private String formatString;
	private final DataFormatter formatter = new DataFormatter();

	private int currentColumn = -1;

	// The last column
	private int lastColumnNumber = -1;
	private boolean bIsFirstRow;
	private int rownum = -1;
	private int lastCellNum;

	private ExcelDTO objExcelDTO = null;

	private Map<Integer, List<ExcelDTO>> map = null;

	private static String inlineStr = "inlineStr";
	private static String cellValue = "v";
	private static String row = "row";
	private static String rowIndexOne = "1";
	private static String cell = "c";
	private static String xmlRow = "r";
	private static String cellTypeInXML = "t";
	private static String cellStyle = "s";
	private static String booleanType = "b";
	private static String errorType = "e";
	private static String stringType = "str";
	private static String xmlDateFormat = "m/d/yy";
	private static String excelDateFormat = "yyyy-MM-dd";
	private static String sstIndex = "s";

	/**
	 * Instantiates a new sheet handler.
	 * 
	 * @param sst
	 *            the sst
	 * @param lastCellNum
	 *            the last cell num
	 * @param stylesTable
	 *            the styles table
	 * @param map
	 *            the map
	 */
	@SuppressWarnings({ "rawtypes", "unchecked" })
	public SheetHandler(ReadOnlySharedStringsTable sst, int lastCellNum,
			StylesTable stylesTable, Map map) {
		sharedStringsTable = sst;
		this.lastCellNum = lastCellNum;
		this.stylesTable = stylesTable;
		this.map = map;

	}

	/**
	 * Column name to index.
	 * 
	 * @param name
	 *            the name
	 * @return the int
	 */
	private int columnNameToIndex(String name) {
		int column = -1;
		for (int i = 0; i < name.length(); ++i) {
			int c = name.charAt(i);
			column = (column + 1) * 26 + c - 'A';
		}
		return column;
	}

	/*
	 * (non-Javadoc)
	 * 
	 * @see org.xml.sax.helpers.DefaultHandler#startElement(java.lang.String,
	 * java.lang.String, java.lang.String, org.xml.sax.Attributes)
	 */
	public void startElement(String uri, String localName, String name,
			Attributes attributes) throws SAXException {
		if (inlineStr.equals(name) || cellValue.equals(name)) {
			vExists = true;
		}
		// row => row
		else if (row.equals(name)) {
			String r = attributes.getValue(xmlRow);
			rownum = Integer.parseInt(r);
			if (rowIndexOne.equals(r)) {
				// indicates first row
				bIsFirstRow = true;
			} else {
				bIsFirstRow = false;
			}
			listExcelDTO = new ArrayList<ExcelDTO>();
		} else if (cell.equals(name)) {
			// Get the cell reference
			// Clear contents cache
			value.setLength(0);

			objExcelDTO = new ExcelDTO();

			String r = attributes.getValue(xmlRow);
			int firstDigit = -1;
			for (int c = 0; c < r.length(); ++c) {
				if (Character.isDigit(r.charAt(c))) {
					firstDigit = c;
					break;
				}
			}

			currentColumn = columnNameToIndex(r.substring(0, firstDigit));

			// Set up defaults.
			this.nextDataType = xssfDataType.NUMBER;
			this.formatIndex = -1;
			this.formatString = null;
			String cellType = attributes.getValue(cellTypeInXML);
			String cellStyleStr = attributes.getValue(cellStyle);

			if (cellType == null && cellStyleStr == null) {
				nextDataType = xssfDataType.NO_STYLE;
			}

			if (booleanType.equals(cellType)) {
				nextDataType = xssfDataType.BOOL;
			} else if (errorType.equals(cellType)) {
				nextDataType = xssfDataType.ERROR;
			} else if (inlineStr.equals(cellType)) {
				nextDataType = xssfDataType.INLINESTR;
			} else if (sstIndex.equals(cellType)) {
				nextDataType = xssfDataType.SSTINDEX;
			} else if (stringType.equals(cellType)) {
				nextDataType = xssfDataType.FORMULA;
			} else if (cellStyleStr != null) {
				// It's a number, but almost certainly one
				// with a special style or format
				int styleIndex = Integer.parseInt(cellStyleStr);
				XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
				this.formatIndex = style.getDataFormat();
				this.formatString = style.getDataFormatString();

				if (this.formatString == null) {
					this.formatString = BuiltinFormats
							.getBuiltinFormat(this.formatIndex);
				}
			}
		}

		if (cell.equals(name) || cellValue.equals(name)) {
			if (currentColumn >= lastCellNum) {
				// stop processing
				return;
			}
		}
	}

	/*
	 * (non-Javadoc)
	 * 
	 * @see org.xml.sax.helpers.DefaultHandler#endElement(java.lang.String,
	 * java.lang.String, java.lang.String)
	 */
	public void endElement(String uri, String localName, String name)
			throws SAXException {

		String thisStr = null;
		if (cell.equals(name) || cellValue.equals(name)) {

			if (currentColumn >= lastCellNum) {
				// stop processing
				return;
			}
		}
		if (cell.equals(name)) {
			if (null == value || "".equals(value.toString())) {
				objExcelDTO.setColumnValue("");
			}

			if (!bIsFirstRow) {
				objExcelDTO.setColumnName(headerList.get(currentColumn));
			}
			listExcelDTO.add(objExcelDTO);
		}

		// v => contents of a cell
		else if (cellValue.equals(name)) {
			// Process the value contents as required.
			// Do now, as characters() may be called more than once
			switch (nextDataType) {

			case BOOL:
				char first = value.charAt(0);
				thisStr = first == '0' ? "FALSE" : "TRUE";
				objExcelDTO.setColumnValue(thisStr);
				break;

			case ERROR:
				thisStr = value.toString();
				objExcelDTO.setColumnValue(thisStr);
				break;

			case FORMULA:
				// A formula could result in a string value,
				// so always add double-quote characters.
				thisStr = value.toString();
				objExcelDTO.setColumnValue(thisStr);
				break;

			case INLINESTR:
				XSSFRichTextString rtsi = new XSSFRichTextString(
						value.toString());
				thisStr = rtsi.toString();
				objExcelDTO.setColumnValue(thisStr);
				break;

			case NO_STYLE:
				XSSFRichTextString its = new XSSFRichTextString(
						value.toString());
				thisStr = its.toString();
				objExcelDTO.setColumnValue(thisStr);
				break;

			case SSTINDEX:
				String sstlndex = value.toString();
				try {
					int idx = Integer.parseInt(sstlndex);
					XSSFRichTextString rtss = new XSSFRichTextString(
							sharedStringsTable.getEntryAt(idx));
					thisStr = rtss.toString();
					objExcelDTO.setColumnValue(thisStr);
				} catch (NumberFormatException ex) {
					LOG.error("Failed to parse SST index '" + sstIndex + "': "
							+ ex.toString());
				}

				break;

			case NUMBER:
				String n = value.toString();
				if (this.formatString != null) {
					if (xmlDateFormat.equals(this.formatString)) {
						this.formatString = excelDateFormat;
					}
					thisStr = formatter.formatRawCellContents(
							Double.parseDouble(n), this.formatIndex,
							this.formatString);
					objExcelDTO.setColumnValue(thisStr);
				} else {
					thisStr = n;
					objExcelDTO.setColumnValue("");
				}
				break;

			default:
				thisStr = "Unexpected type: " + nextDataType;
				break;
			}

			if (bIsFirstRow) {
				objExcelDTO.setColumnName("HEADER");
				headerList.add(thisStr);
			}

			if (lastColumnNumber == -1) {
				lastColumnNumber = 0;
			}

			if (null == thisStr || "".equals(thisStr)) {
				objExcelDTO.setColumnValue("");
			}

			// Update column
			if (currentColumn > -1) {
				lastColumnNumber = currentColumn;
			}
			vExists = false;
		} else if (row.equals(name)) {
			if (minColumnCount > 0) {
				// Columns are 0 based
				if (lastColumnNumber == -1) {
					lastColumnNumber = 0;
				}
			}

			// We're onto a new row
			lastColumnNumber = -1;
			map.put(rownum - 1, listExcelDTO);
		}
	}

	/*
	 * (non-Javadoc)
	 * 
	 * @see org.xml.sax.helpers.DefaultHandler#characters(char[], int, int)
	 */
	public void characters(char[] ch, int start, int length)
			throws SAXException {

		if (vExists) {
			value.append(ch, start, length);
		}
	}
}
