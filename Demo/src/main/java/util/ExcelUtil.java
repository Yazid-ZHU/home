package util;

import java.io.File;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;



public class ExcelUtil {

	//@Test
	public void testExelContent() throws Exception {
		// XSSFWorkbook, File
		OPCPackage pkg = OPCPackage.open(new File("D:\\BillFiles\\test\\test.xlsx"));
		XSSFWorkbook wb = new XSSFWorkbook(pkg);

		DataFormatter formatter = new DataFormatter();
		Sheet sheet = wb.getSheetAt(0);

		// Decide which rows to process
		int rowStart = Math.min(0, sheet.getFirstRowNum());
		int rowEnd = Math.max(20, sheet.getLastRowNum());

		for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
			Row r = sheet.getRow(rowNum);
			if (r == null) {
				// This whole row is empty
				// Handle it as needed
				continue;
			}

			int MY_MINIMUM_COLUMN_COUNT = 3;
			int lastColumn = Math.max(r.getLastCellNum(), MY_MINIMUM_COLUMN_COUNT);

			for (int cn = 0; cn < lastColumn; cn++) {
				Cell cell = r.getCell(cn, Row.CREATE_NULL_AS_BLANK);
				int type = cell.getCellType();

				switch (type) {
				case 1:
					System.out.println(cell.getRichStringCellValue().getString());
					break;
				case 0:
					if (DateUtil.isCellDateFormatted(cell)) {
						System.out.println(cell.getDateCellValue());
					} else {
						System.out.println(cell.getNumericCellValue());
					}
					break;
				case 4:
					System.out.println(cell.getBooleanCellValue());
					break;
				case 2:
					System.out.println(cell.getCellFormula());
					break;
				case 3:
					System.out.println("@");
					break;
				default:
					System.out.println();
				}
			}
		}

		pkg.close();
	}

	//@Test
	public void testSaxProcessor() throws Exception {
		processOneSheet("D:\\BillFiles\\test\\11111.xlsx");
	}

	public void processOneSheet(String filename) throws Exception {
		OPCPackage pkg = OPCPackage.open(filename);
		XSSFReader reader = new XSSFReader(pkg);
		SharedStringsTable sst = reader.getSharedStringsTable();

		// 1.获取读取器
		XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");

		// 2.设置内容处理器
		SheetHandler handler = new SheetHandler(sst);
		parser.setContentHandler(handler);

		// 3.读取xml文件內容
		InputStream sheet1 = reader.getSheet("rId1");
		InputSource sheetSource = new InputSource(sheet1);
		parser.parse(sheetSource);
		sheet1.close();

		// 4.从内容处理器中获取返回的list
		List<MRow> rowList = handler.getRowList();
		int rowLenth = rowList.get(0).getCells().size();
		System.out.println("rowLenth: " + rowLenth);

		// 5.重置表格内容
		for (MRow row : rowList) {
			int rowId = row.getRowNum();
			System.out.println("row number: " + rowId);
			
			Map<String, MCell> cells = row.getCells();

			Map<String, MCell> newCells = new LinkedHashMap<String, MCell>();

			for (int cn = 1; cn < rowLenth + 1; cn++) {
				MCell ncell = new MCell();

				String cellKey = rowId + "" + cn;
				if (cells.get(cellKey) == null) {
					newCells.put(cellKey, new MCell(Integer.parseInt(cellKey), "@"));
				} else {
					newCells.put(cellKey, cells.get(cellKey));
				}
			}

			row.setCells(newCells);
		}

		 //6.打印检测表格
		 for(MRow row : rowList){
			 System.out.println("row number: " + row.getRowNum());
			 Map<String, MCell> cells = row.getCells();
			
			 for(Entry<String, MCell> entry : cells.entrySet()){
			 System.out.println(" map key: " + entry.getKey());
			 System.out.println(" cell value: " + entry.getValue());
			 }
		 }
	}

	// public void processAllSheets(String filename) throws Exception {
	// OPCPackage pkg = OPCPackage.open(filename);
	// XSSFReader reader = new XSSFReader( pkg );
	// SharedStringsTable sst = reader.getSharedStringsTable();
	//
	// XMLReader parser = fetchSheetParser(sst);
	//
	// Iterator<InputStream> sheets = reader.getSheetsData();
	// while(sheets.hasNext()) {
	// System.out.println("Processing new sheet:\n");
	// InputStream sheet = sheets.next();
	// InputSource sheetSource = new InputSource(sheet);
	// parser.parse(sheetSource);
	// sheet.close();
	// System.out.println("");
	// }
	// }

	// public XMLReader fetchSheetParser(SharedStringsTable sst) throws
	// SAXException {
	// XMLReader parser =
	// XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
	// ContentHandler handler = new SheetHandler(sst);
	// parser.setContentHandler(handler);
	// return parser;
	// }

	/**
	 * 内容处理器
	 */
	private class SheetHandler extends DefaultHandler {
		private SharedStringsTable sst;
		private String cellValue;
		private boolean isIndex;

		private List<MRow> rowList = new ArrayList<MRow>();
		private MRow mrow;
		private MCell mcell;

		public List<MRow> getRowList() {
			return rowList;
		}

		private SheetHandler(SharedStringsTable sst) {
			this.sst = sst;
		}

		public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {

			if ("row".equals(name)) {
				mrow = new MRow();
				int rowId = Integer.parseInt(attributes.getValue("r"));
				mrow.setRowNum(rowId);
			}

			// c => cell
			if ("c".equals(name)) {
				mcell = new MCell();
				// get cell position
				String cellPosition = attributes.getValue("r");
				int rowNum = DigitUtil.getNumbers(cellPosition);
				int columnNum = DigitUtil.charToNum(cellPosition.charAt(0));
				int cellNum = Integer.parseInt(rowNum + "" + columnNum);
				mcell.setCellNum(cellNum);

				// Figure out if the value is an index in the SST
				String cellType = attributes.getValue("t");
				if (cellType != null && cellType.equals("s")) {// 要去查询sst
					isIndex = true;
				} else {
					isIndex = false;
				}
			}
			// Clear contents cache
			cellValue = "";
		}

		public void characters(char[] ch, int start, int length) throws SAXException {
			cellValue += new String(ch, start, length);
		}

		public void endElement(String uri, String localName, String name) throws SAXException {
			// Process the last contents as required.
			// Do now, as characters() may be called more than once
			if (isIndex) {
				int idx = Integer.parseInt(cellValue);
				cellValue = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
				isIndex = false;

			}

			// v => contents of a cell
			if (name.equals("v")) {
				mcell.setValue(cellValue);
				mrow.getCells().put(mcell.getCellNum()+"", mcell);
			}

			if (name.equals("c")) {
				mcell = null;
			}

			if (name.equals("row")) {
				rowList.add(mrow);
				mrow = null;
			}
		}
	}

}
