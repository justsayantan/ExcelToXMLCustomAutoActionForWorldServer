package com.idiominc.ws.component.autoaction;

import java.io.File;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import com.idiominc.wssdk.WSContext;
import com.idiominc.wssdk.WSException;
import com.idiominc.wssdk.ais.WSNode;
import com.idiominc.wssdk.asset.WSAssetTask;
import com.idiominc.wssdk.workflow.WSProject;
import com.idiominc.wssdk.workflow.WSTask;
import com.idiominc.wssdk.workflow.WSWorkflow;
import com.idiominc.wssdk.component.autoaction.WSActionResult;
import com.idiominc.wssdk.component.autoaction.WSTaskAutomaticAction;

import org.apache.commons.io.FileUtils;
import org.apache.log4j.Category;
import org.apache.log4j.Level;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class ExcelToXml extends WSTaskAutomaticAction {

	private static final String DONE_RETURN_VALUE = "Done";
	@SuppressWarnings("deprecation")
	private static final Category log = Category.getInstance(ExcelToXml.class.getName());
	private static final int PER_SIZE = 200;
	private WSProject project = null;

	@Override
	public String getDescription() {
		// TODO Auto-generated method stub
		String AUTOACTION_DESC = "ExcelToXml";
		return (AUTOACTION_DESC);

	}

	@Override
	public String getName() {
		// TODO Auto-generated method stub
		String AUTOACTION_NAME = "ExcelToXml";
		return (AUTOACTION_NAME);
	}

	@Override
	public String[] getReturns() {
		// TODO Auto-generated method stub

		String[] AUTOACTION_RETURN_VALUES = new String[] { DONE_RETURN_VALUE };

		return (AUTOACTION_RETURN_VALUES);
	}

	@Override
	public WSActionResult execute(WSContext context, Map parameters, WSTask task) throws WSException {
		log.setLevel(Level.DEBUG);
		WSAssetTask assetTask = (WSAssetTask) task;
		if (assetTask.getSourceAisNode().getName().contains(".xlsx")) {
			// Get the Temp folder path from Configuration
			File tempFolder = context.getConfigurationManager().getTemporaryDirectory();
			log.debug("Temp = " + tempFolder.getPath());

			// Declare Variables
			ArrayList<WSNode> nodeList = new ArrayList<WSNode>();
			Workbook book = null;
			String path = null;
			File xmlPath = null;
			String xmlPathString = null;
			ArrayList<ArrayList<String>> data = new ArrayList<ArrayList<String>>();
			ArrayList<String> xmlSourcePaths = new ArrayList<String>();

			// Get the source folder in a project
			path = assetTask.getSourceAisNode().getFile().getParent();
			
			// Get the path of the generated XML file in the temp.
			xmlPath = new File(tempFolder, assetTask.getSourceAisNode().getName());
			log.debug("SourceNode Path : " + path);
			log.debug("XML Directory Path :" + xmlPath.getPath());

			// Get the Source Node(Excel File Node)
			WSNode excelNode = assetTask.getSourceAisNode();
			try {
				// Read the file and assigned it to the book
				book = (Workbook) WorkbookFactory.create(excelNode.getFile());
				log.debug("Book Created ");
			} catch (Exception e) {
				throw new WSException(e);
			}

			// Initializing the XML document
			DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
			DocumentBuilder builder = null;
			try {
				builder = factory.newDocumentBuilder();
			} catch (ParserConfigurationException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			Document document = builder.newDocument();
			
			Sheet sheet = book.getSheetAt(0); // Read the first sheet
			int totalRowCount = sheet.getLastRowNum(); //Get the total rows count 
			log.debug("Sheet = " + sheet.getSheetName());
			
			Iterator<?> rows = sheet.rowIterator(); // Get Rows
			
			//Implemented the logic to split large excel file to multiple xml file
			int count = 0;
			int next = 0;
			if (PER_SIZE > totalRowCount) {
				next = totalRowCount;
			} else {
				next = PER_SIZE;
			}

			while (rows.hasNext()) {
				Row row = (Row) rows.next();
				if (count == next) {
					try {
						// replacing the path with .xml
						xmlPathString = xmlPath.getPath().replace(".xlsx", count + ".xml");
						log.debug("XML File Path :" + xmlPathString);
						String pathGenerated = SaveXml(data, xmlPathString);//Saving the data
						xmlSourcePaths.add(pathGenerated);
						if ((PER_SIZE + next) > totalRowCount) {
							next = totalRowCount;
						} else {
							next = PER_SIZE + next;
						}
					} catch (UnsupportedEncodingException e) {
						// TODO Auto-generated catch block
						log.error("Exception" + e.getMessage());
					} catch (ParserConfigurationException e) {
						// TODO Auto-generated catch block
						log.error("Exception" + e.getMessage());
					} catch (TransformerException e) {
						// TODO Auto-generated catch block
						log.error("Exception" + e.getMessage());
					}
					ArrayList<String> firstRowData = data.get(0);
					data = new ArrayList<ArrayList<String>>();
					data.add(firstRowData);
				}

				ArrayList<String> rowData = processOneRow(row, document);
				data.add(rowData);
				count++;
			}

			try {
				book.close();
				data.clear();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}

			project = task.getProject();

			// logic to get list of WSNode
			if (xmlSourcePaths != null && xmlSourcePaths.size() > 0) {
				for (String xmlSourcePath : xmlSourcePaths) {
					log.debug("XML Source Path : " + xmlSourcePath);

					File source = new File(xmlSourcePath);
					log.debug("Source Path : " + source.getPath());
					File dest = new File(path, source.getName());
					log.debug("Destination Path : " + dest.getPath());
					try {
						FileUtils.moveFile(source, dest);
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					String xmlNodePath = assetTask.getSourceAisNode().getParent().getPath();
					xmlNodePath = xmlNodePath + "/" + source.getName();
					log.debug("XMLNode path: " + xmlNodePath);
					WSNode node = context.getAisManager().getNode(xmlNodePath);
					nodeList.add(node);
				}
			}

			WSNode[] nodes = nodeList.toArray(new WSNode[0]);
			WSWorkflow wf = task.getWorkflow();
			WSTask[] xmlTask = project.createTasks(nodes, wf);//Create the tasks
			project.addTasks(xmlTask);//Add the tasks into project
			log.debug("New Task Created : ");

			return (new WSActionResult(DONE_RETURN_VALUE, "New Task added with XML file"));
		}
		return null;
	}

	//Method to Save Data into XML.
	private String SaveXml(ArrayList<ArrayList<String>> data, String path)
			throws ParserConfigurationException, TransformerException, UnsupportedEncodingException {
		// TODO Auto-generated method stub

		DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
		DocumentBuilder builder = factory.newDocumentBuilder();
		Document document = builder.newDocument();
		Element rootElement = document.createElement("rows");
		document.appendChild(rootElement);

		int numOfProduct = data.size();

		for (int i = 1; i < numOfProduct; i++) {
			Element productElement = document.createElement("row");
			rootElement.appendChild(productElement);

			int index = 0;
			for (String s : data.get(i)) {
				String headerString = data.get(0).get(index);
				if (headerString.contains(" ")) {
					headerString = headerString.replace(" ", "_");
				}
				byte[] byteText = headerString.getBytes(Charset.forName("UTF-8"));
				// To get original string from byte.
				String originalString = new String(byteText, "UTF-8");
				Element headerElement = document.createElement(originalString);
				productElement.appendChild(headerElement);
				headerElement.appendChild(document.createTextNode(s));
				index++;
			}
		}

		// Write to the XML file
		TransformerFactory tFactory = TransformerFactory.newInstance();
		Transformer transformer = tFactory.newTransformer();

		// Add indentation to output
		transformer.setOutputProperty(OutputKeys.INDENT, "yes");
		transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");

		DOMSource source = new DOMSource(document);

		String xmlPath = path;
		File file = new File(xmlPath);
		log.debug("New File Path: " + file.getPath());
		StreamResult result = new StreamResult(file.getPath());
		log.debug("Generated XML : " + result);
		transformer.transform(source, result);

		return xmlPath;

	}

	//Method to read Data For One Row.
	public static ArrayList<String> processOneRow(Row row, Document xml) {
		try {

			int rowNumber = row.getRowNum();
			// display row number
			System.out.println("Row No.: " + rowNumber);

			// get a row, iterate through cells.
			Iterator<?> cells = row.cellIterator();

			ArrayList<String> rowData = new ArrayList<String>();
			while (cells.hasNext()) {
				XSSFCell cell = (XSSFCell) cells.next();
				// System.out.println ("Cell : " + cell.getCellNum ());
				switch (cell.getCellType()) {
				case XSSFCell.CELL_TYPE_NUMERIC: {
					// NUMERIC CELL TYPE
					System.out.println("Numeric: " + cell.getNumericCellValue());
					rowData.add(cell.getNumericCellValue() + "");
					break;
				}
				case XSSFCell.CELL_TYPE_STRING: {
					// STRING CELL TYPE
					XSSFRichTextString richTextString = cell.getRichStringCellValue();
					byte[] byteText = richTextString.toString().getBytes(Charset.forName("UTF-8"));
					// To get original string from byte.
					String originalString = new String(byteText, "UTF-8");

					System.out.println("String: " + new String(originalString.getBytes("UTF-8")));
					rowData.add(originalString);
					break;
				}
				default: {
					// types other than String and Numeric.
					System.out.println("Type not supported.");
					break;
				}
				} // end switch

			} // end while
			return rowData;
		} catch (IOException e) {
			log.error("IOException " + e.getMessage());
		}
		return null;
	}

}
