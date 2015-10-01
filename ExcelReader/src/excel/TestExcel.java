package excel;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.xml.sax.SAXException;

/**
 * The Class TestExcel.
 * 
 * @author Sourav
 */
public class TestExcel {

	/**
	 * The main method.
	 *
	 * @param args the arguments
	 * @throws IOException Signals that an I/O exception has occurred.
	 * @throws SAXException the SAX exception
	 * @throws OpenXML4JException the open xm l4 j exception
	 */
	public static void main(String[] args) throws IOException, SAXException,
			OpenXML4JException {

		File excelFile = new File("C:\\Users\\User\\Desktop\\ExcelReader\\Test.xls");
		Map<Integer, List<ExcelDTO>> excelMap = ExcelParser.read(excelFile, 4, 1);

		for (Map.Entry<Integer, List<ExcelDTO>> entry : excelMap.entrySet()) {

			System.out.println("Row Number " + entry.getKey());
			for (ExcelDTO dto : entry.getValue()) {
				System.out.println("Column Name: " + dto.getColumnName()
						+ " , Column Value: " + dto.getColumnValue() + ",");
			}
		}
	}

}
