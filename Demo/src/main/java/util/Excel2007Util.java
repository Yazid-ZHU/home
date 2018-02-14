package util;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.logging.Logger;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;


public class Excel2007Util {
	//日志
	private static final Logger LOGGER = Logger.getAnonymousLogger("Excel2007Util");
	/**
	 * processOneSheet
	 * @param filename
	 * @return map
	 */
	public Map<String, Object> processOneSheet(String filename){
		// 1.获取读取器
		XMLReader parser = null;
		// 2.设置内容处理器
		SheetHandler handler = null;
		// 3.读取xml文件內容
		InputStream sheet1 = null;
		try {
			OPCPackage pkg = OPCPackage.open(filename);
			XSSFReader reader = new XSSFReader(pkg);
			SharedStringsTable sst = reader.getSharedStringsTable();
			parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");

			handler = new SheetHandler(sst);
			parser.setContentHandler(handler);
			sheet1 = reader.getSheet("rId1");
		} catch (InvalidFormatException e) {
			LOGGER.info("");
		} catch (IOException e) {
			LOGGER.info("IOException error");
		} catch (OpenXML4JException e) {
			LOGGER.info("OpenXML4JException error");
		} catch (SAXException e) {
			LOGGER.info("SAXException error");
		}
		try {
			InputSource sheetSource = new InputSource(sheet1);
			parser.parse(sheetSource);
		} catch (IOException e) {
			LOGGER.info("parser IOException error");
		} catch (SAXException e) {
			LOGGER.info("parser SAXException error");
		}finally {
			if(sheet1 != null){
				try {
					sheet1.close();
				} catch (IOException e) {
					LOGGER.info("close InputStream error");
				}
			}
			
		}

		// 4.从内容处理器中获取返回的list
		List<MRow> rowList = handler.getRowList();
		int rowLenth = rowList.get(0).getCells().size();

		// 5.重置表格内容
		for (MRow row : rowList) {
			int rowId = row.getRowNum();
			//System.out.println("row number: " + rowId);
			
			Map<String, MCell> cells = row.getCells();

			Map<String, MCell> newCells = new LinkedHashMap<String, MCell>();

			for (int cn = 1; cn < rowLenth + 1; cn++) {
				String cellKey = Integer.toString(rowId) + cn;
				if (cells.get(cellKey) == null) {
					newCells.put(cellKey, new MCell(Integer.parseInt(cellKey), ""));
				} else {
					newCells.put(cellKey, cells.get(cellKey));
				}
			}

			row.setCells(newCells);
		}
		Map<String, Object> title = new HashMap<String, Object>();
		List<Map<String, Object>> result = new ArrayList<Map<String, Object>>();
		for(int i=0; i<rowList.size(); i++){
			Integer index = 0;
			Map<String, Object> map = new HashMap<String, Object>();
			Map<String, MCell> cells = rowList.get(i).getCells();
			if(i==0){
				for(Entry<String, MCell> entry : cells.entrySet()){
					title.put(index.toString(), entry.getValue().getValue());
					index++;
				}
			}else{
				for(Entry<String, MCell> entry : cells.entrySet()){
					map.put((String) title.get(index.toString()), entry.getValue().getValue());
					index++;
				}
				result.add(map);
			}
		}
		Map<String, Object> mapR = new HashMap<String, Object>();
        mapR.put("result", result);
        StringBuilder titleList = new StringBuilder();
        for (int i=0; i<title.size();i++) {
        	titleList.append(title.get(Integer.toString(i)));
        	titleList.append(",");
		}
        titleList.substring(0,titleList.length()-1);
        String[] split = titleList.toString().split(",");
        mapR.put("title", Arrays.asList(split));
        if(null==result || result.isEmpty())
        {
        	mapR.put("total", 0);
        }
        else
        {
        	mapR.put("total", result.size());
        }
        mapR.put("errorList", new ArrayList<String>());
        return mapR;
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
		/**
		 * startElement 处理
		 */
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
				int cellNum = Integer.parseInt(Integer.toString(rowNum) + Integer.toString(columnNum));
				mcell.setCellNum(cellNum);

				// Figure out if the value is an index in the SST
				String cellType = attributes.getValue("t");
				if ( null != cellType && "s".equals(cellType)) {// 要去查询sst
					isIndex = true;
				} else {
					isIndex = false;
				}
			}
			// Clear contents cache
			cellValue = "";
		}
		/**
		 * characters 处理
		 */
		public void characters(char[] ch, int start, int length) throws SAXException {
			cellValue += new String(ch, start, length);
		}
		/**
		 * endElement 处理
		 */
		public void endElement(String uri, String localName, String name) throws SAXException {
			if (isIndex) {
				int idx = Integer.parseInt(cellValue);
				cellValue = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
				isIndex = false;

			}

			if ("v".equals(name)) {
				mcell.setValue(cellValue);
				mrow.getCells().put(mcell.getCellNum()+"", mcell);
			}

			if ("c".equals(name)) {
				mcell = null;
			}

			if ("row".equals(name)) {
				rowList.add(mrow);
				mrow = null;
			}
		}
	}

}
