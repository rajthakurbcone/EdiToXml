import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MainProgram {

	public static void main(String[] args) throws IOException, Exception {
		findRelationship("map1.xml", "map2.xml");

	}

	public static HashMap<String, String> findNodeByColumnNumber(String xmlFile) throws Exception {

		System.out.println("Processing XML File...");
		System.out.println("xml File: " + xmlFile);
		HashMap<String, String> keyPairs = new HashMap<String, String>();
		File file = new File(xmlFile);
		DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
		Document doc = dBuilder.parse(file);

		doc.getDocumentElement().normalize();

		NodeList childNodes = doc.getElementsByTagName("*");

		System.out.println();
		for (int i = 0; i < childNodes.getLength(); i++) {

			Node node = childNodes.item(i);

			String nodeName = node.getNodeName();
			if (!nodeName.equals("#text")) {
				Boolean isGroup = node.getTextContent().contains("\n");

				if (!isGroup) {
					String column = node.getTextContent();
//						System.out.println(nodeName+"  "+column);
					keyPairs.put(column, nodeName);

				}
			}
		}
		System.out.println("Total Columns Info Added: " + keyPairs.size());
		System.out.println("..............................");

		return keyPairs;

	}

	public static void findRelationship(String ediFile, String xmlFile) throws Exception {

		HashMap<String, String> findNodeByColumnNumber = findNodeByColumnNumber(xmlFile);
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Mapping");

		int rowCount = 0;
		Row row = sheet.createRow(rowCount);
		int columnCount = -1;
		Cell cell = row.createCell(++columnCount);
		cell.setCellValue((String) "EDI Field");
		cell = row.createCell(++columnCount);
		cell.setCellValue((String) "Name");
		 cell = row.createCell(++columnCount);
		cell.setCellValue((String) "Column");
		 cell = row.createCell(++columnCount);
		cell.setCellValue((String) "XML Field");

		File file = new File(ediFile);
		DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
		Document doc = dBuilder.parse(file);

		doc.getDocumentElement().normalize();

		NodeList childNodes = doc.getDocumentElement().getChildNodes();

		for (int i = 0; i < childNodes.getLength(); i++) {
			try {
				String nodeName = childNodes.item(i).getNodeName();
				if (!nodeName.equals("#text")) {
					String node = nodeName;
					String nameAttr = childNodes.item(i).getAttributes().getNamedItem("name").getNodeValue();
					String columnAttr = childNodes.item(i).getAttributes().getNamedItem("column").getNodeValue();
					String fieldInXml = findNodeByColumnNumber.get("ASCIICOL_" + columnAttr);
					System.out.println("EDI: " + nodeName);
					System.out.println("Name: " + nameAttr);
					System.out.println("Column: " + columnAttr);
					System.out.println("In XSL: " + fieldInXml);
					System.out.println("----------------------------");
					 row = sheet.createRow(++rowCount);
					 columnCount = -1;
					 cell = row.createCell(++columnCount);
					cell.setCellValue((String) nodeName);
					cell = row.createCell(++columnCount);
					cell.setCellValue((String) nameAttr);
					 cell = row.createCell(++columnCount);
					cell.setCellValue((String) columnAttr);
					 cell = row.createCell(++columnCount);
					cell.setCellValue((String) fieldInXml);
				}
			} catch (Exception e) {
				// TODO: handle exception
			}
		}

		FileOutputStream outputStream = new FileOutputStream("MappingSheet.xlsx");
		workbook.write(outputStream);

	}
}
