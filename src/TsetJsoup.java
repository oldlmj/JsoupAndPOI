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

public class TsetJsoup {
	/**
	 * Jsoup 爬蟲 因原本的不見，所以重做一份
	 * - 說明
	 * 1.依據該 Excel 裡的所有商品來爬取所需資訊
	 * 2.而所需資訊為:商品描述、在"了解更多"下方全部圖片的 url
	 * 3.再回寫到該 Excel 指定欄位裡。
	 * 
	 * - 使用 Jsoup、Apache POI。
	 * 
	 * log 
	 * ***** 201118
	 * - 已抓到商品描述
	 * ***** 201119
	 * - 已抓到圖片的src 
	 * - 可讀寫 Excel 檔案內容 
	 * ****** 201120
	 * - 新增 GrabDataToExcel.java，製造商品資料 Excel,
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 * 
	 */

	// static String[] strData = new String[2];

	public static String[] getData(String str) {
		String strUrl = "https://www.woodstuck.com.tw/products/" + str;
		String strTmp = "";
		System.out.println(strUrl + "----");
		String[] strData = new String[2]; //放置爬蟲捕獲的資料
		try {

			Document doc = Jsoup.connect(strUrl).timeout(30000).validateTLSCertificates(false).get();
			// ## 取得商品描述
			Elements el = doc.select("p.MsoNormal");
			for (Element postItem : el) {
				// 因 Shopline 的系統可接受HTML語法，因此在每一行加入<br>來換行
				strTmp = strTmp + postItem.text() + "<br>";
			}

			strData[0] = strTmp;			
			strTmp = "";
			
			// ## 取得所有圖片 url			
			Elements elDescription = doc.select("p.text-center");
			for (Element elpostItem : elDescription) {
				Elements elUrlItem = elpostItem.getElementsByTag("img");
				// 依據 Shopline 規則，若要放多個圖片連結，在連結後面再放個半行空白即可
				strTmp = strTmp + elUrlItem.attr("src") + "　";

			}
			// 把 strTmp 放進字串陣列
			strData[1] = strTmp;			
		} catch (IOException e) {
			e.printStackTrace();
		}
		return strData;
	}
	// 讀取跟寫入 Excel 內容
	public static void readAndWriteExcel() throws FileNotFoundException, IOException {		
		String fileName = "test_ws_data.xlsx";
		String[] str = new String[2];
		
		try {
			FileInputStream input = new FileInputStream(fileName); // 輸入串流			
			XSSFWorkbook book = new XSSFWorkbook(input); //建立活頁簿

			XSSFSheet sheet = book.getSheetAt(0); // 選擇第一個工作表
			
			for (int i = 1; i <= sheet.getLastRowNum(); i++) {
				// 得到列
				XSSFRow row = sheet.getRow(i);				
				// 在第2行依序讀取儲存格，依序放進 getData( String str ) 提供爬蟲捕抓的目標 
				
				try {
					str = getData(row.getCell(1).getStringCellValue());                                    
                } catch (NullPointerException e) {
                    //如果儲存格為空，就跳過此循環
                    continue;
                }
				// 依序選取指定的儲存格，並放置 str[] 的每個值
				XSSFCell cell = row.createCell(3);
				cell.setCellValue(str[0]);
				cell = row.createCell(7);
				cell.setCellValue(str[1]);
			}
			input.close(); // 輸入串流關閉
			FileOutputStream out = new FileOutputStream(fileName); // 輸出串流
			book.write(out);
			out.flush(); // 強制將緩衝區的資料輸出
			if (out != null) {
				out.close(); // 輸出串流關閉
			}
			book.close(); // 關閉活頁簿
			System.out.println(fileName + " excel export finish. -------------");
		} catch (FileNotFoundException e) {
			System.err.println("OOPS! 檔案不存在~!" + e.toString());
		} catch (IOException e) {
			System.err.println("OOPS! 檔案處理出問題了~!" + e.toString());
		} catch (Exception e) {
			System.err.println("OOPS!!問題可不小..." + e.toString());
		}

	}

	public static void main(String[] args) throws FileNotFoundException, IOException {
		readAndWriteExcel();
	}


}
