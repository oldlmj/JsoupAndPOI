import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.security.SecureRandom;
import java.security.cert.CertificateException;
import java.security.cert.X509Certificate;
import java.util.ArrayList;
import java.util.List;

import javax.net.ssl.HostnameVerifier;
import javax.net.ssl.HttpsURLConnection;
import javax.net.ssl.SSLContext;
import javax.net.ssl.SSLSession;
import javax.net.ssl.X509TrustManager;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

public class GrabDataToExcel {
	/**
	 * 前言: 因原本的爬蟲專案沒備份到，要重做一份。
	 * 
	 * 說明: 當時新增的商品，幾乎完售被下架， 使得無法使用當時所用商品資料的Excel檔進行爬蟲動作， 因此要做個小爬蟲，只抓個商品名稱，並寫入Excel，
	 * 來取代原本商品資料的Excel。
	 * 
	 */
	public static void main(String[] args) throws FileNotFoundException, IOException {
		writeExcel();
	}

	// 讀取 Excel 內容
	public static void writeExcel() throws FileNotFoundException, IOException {
		@SuppressWarnings("resource")
		// 新建工作簿
		XSSFWorkbook book = new XSSFWorkbook();
		// 建立工作表
		XSSFSheet sheet = book.createSheet("data");

		String[] buffer = { "商品編號", "商品名稱(中文)", "商品名稱 (英文)", "商品描述(中文)", "商品描述 (英文)", "SEO標題(中文)", "SEO標題(英文)", "商品相片",
				"更多相片", "選項貨號" };

		int rowIdx = 0;
		int colIdx = -1;

		// 先將 String[] buffer 依序寫入 Excel 的第1列為標題列
		XSSFRow row = sheet.createRow(0); // 建立行
		XSSFCell cell = row.createCell(1);
		for (String arrs : buffer) {
			cell = row.createCell(++colIdx);
			cell.setCellValue(arrs);
			
		}
		// 從 getData() 回傳的字串陣列，在 Excel 的(B,2)開始依序垂直寫入
		String[] strData = getData();
		for (String arrs : strData) {
			row = sheet.createRow(++rowIdx);
			cell = row.createCell(1);
			cell.setCellValue(arrs);
			sheet.autoSizeColumn(rowIdx); // 自動調整欄位寬度
		}
		
		// 指定檔案名稱
		String fileName = "test_ws_data.xlsx";

		/*
		 * 未指定檔案路徑，檔案建立在本執行專案內 儲存工作簿
		 */
		try {
			FileOutputStream os = new FileOutputStream(fileName);
			book.write(os);
			book.close();
			System.out.println(fileName + " excel export finish.");
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	// 使用 Jsoup 爬取商品名稱
	public static String[] getData() {
		String strUrl = "https://www.woodstuck.com.tw/categories/%E5%A4%96%E5%A5%97-1";
		String str = null;
		try {
			Document doc = Jsoup.connect(strUrl).timeout(30000).validateTLSCertificates(false).get();

			Elements el1 = doc.select("ul.boxify-container");
			// ## 取得該頁面所有商品名稱
			for (Element postItem : el1) {
				Elements el = postItem.select("div.title");
				str = el.text();
				System.out.println(el.text());
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
		// str 的內容是"商品名稱"+" "的一直連續性的字串，因此要做切割，並放進 String[] data 裡
		String[] data = str.split(" ");
		return data;
	}

}
