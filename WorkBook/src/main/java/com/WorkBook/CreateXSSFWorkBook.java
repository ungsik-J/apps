package com.WorkBook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.Properties;
import java.util.regex.Matcher;

import org.apache.commons.lang3.StringUtils;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

public class CreateXSSFWorkBook {
	private static final Logger LOG = Logger.getLogger(CreateXSSFWorkBook.class);

	private static String Path = "D:\\DevHome\\git\\apps\\WorkBook\\src\\attachmentFile\\";

	private static String OsFilePath = Path.replaceAll("/", Matcher.quoteReplacement(File.separator));
	private static String ReverseSlashPath = OsFilePath.replaceAll(Matcher.quoteReplacement(File.separator), "/");

	private static String TemplateXlsx = ReverseSlashPath + "/source/Template.xlsx";
	private static String TemplateJSON = ReverseSlashPath + "/source/Template.json";

	private static String CreateFilePath = ReverseSlashPath + "/download/";

	private static File SheetCellProperties = new File(
			"D:\\DevHome\\git\\apps\\WorkBook\\src\\main\\resource\\SheetCellColumn.properties");

	private static long CreateRow = 0L;

	public static void main(String[] args) {
		LOG.info("CreateXSSFWorkBook");

		JSONParser parser = new JSONParser();
		Object pObj = null;
		try {
			pObj = parser.parse(new FileReader(TemplateJSON));
		} catch (FileNotFoundException e) {
			LOG.error("★Run.FileNotFoundException", e);
		} catch (IOException e) {
			LOG.error("★Run.IOException", e);
		} catch (ParseException e) {
			LOG.error("★Run.ParseException", e);
		}
		JSONObject JSONObj = (JSONObject) pObj;
		JSONArray ItemList = (JSONArray) JSONObj.get("NEW_CREATE_FILE_NAME");
		for (Object ItemObj : ItemList) {
			Run(ItemObj);
		}
	}

	@SuppressWarnings("unused")
	public static void Run(Object obj) {

		String Item[] = new String[] { String.valueOf(obj) };

		String IfId = Item[0];

		Workbook wb = null;
		try {
			wb = new XSSFWorkbook(new FileInputStream(TemplateXlsx));
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		int phx, phy = 0;
		String phys, getCellValue = null;
		Sheet sheet = null;
		Properties prop = null;
		for (int i = 0; i < wb.getNumberOfSheets(); i++) {
			sheet = wb.getSheetAt(i);
			for (Row row : sheet) {
				for (Cell cell : row) {
					phx = cell.getColumnIndex();
					phy = cell.getRowIndex();
					phys = (phx + "." + phy);
					try (FileReader file = new FileReader(SheetCellProperties)) {
						prop = new Properties();
						prop.load(file);
						getCellValue = cell.getStringCellValue();
					} catch (Exception e) {
						LOG.error(e);
					}

				}
			}
		}
		String FileName = String.valueOf(obj);
		try {
			FileOutputStream fileOut = new FileOutputStream(new File(CreateFilePath + FileName + ".xlsx"));
			try {
				wb.write(fileOut);
			} catch (IOException e) {
				e.printStackTrace();
			} finally {
				try {
					fileOut.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} finally {
			CreateRow += 1;
		}
	}
}
