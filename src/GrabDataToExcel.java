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
	 * �e��: �]�쥻�����αM�רS�ƥ���A�n�����@���C
	 * 
	 * ����: ��ɷs�W���ӫ~�A�X�G����Q�U�[�A �ϱo�L�k�ϥη�ɩҥΰӫ~��ƪ�Excel�ɶi�檦�ΰʧ@�A �]���n���Ӥp���ΡA�u��Ӱӫ~�W�١A�üg�JExcel�A
	 * �Ө��N�쥻�ӫ~��ƪ�Excel�C
	 * 
	 */
	public static void main(String[] args) throws FileNotFoundException, IOException {
		writeExcel();
	}

	// Ū�� Excel ���e
	public static void writeExcel() throws FileNotFoundException, IOException {
		@SuppressWarnings("resource")
		// �s�ؤu�@ï
		XSSFWorkbook book = new XSSFWorkbook();
		// �إߤu�@��
		XSSFSheet sheet = book.createSheet("data");

		String[] buffer = { "�ӫ~�s��", "�ӫ~�W��(����)", "�ӫ~�W�� (�^��)", "�ӫ~�y�z(����)", "�ӫ~�y�z (�^��)", "SEO���D(����)", "SEO���D(�^��)", "�ӫ~�ۤ�",
				"��h�ۤ�", "�ﶵ�f��" };

		int rowIdx = 0;
		int colIdx = -1;

		// ���N String[] buffer �̧Ǽg�J Excel ����1�C�����D�C
		XSSFRow row = sheet.createRow(0); // �إߦ�
		XSSFCell cell = row.createCell(1);
		for (String arrs : buffer) {
			cell = row.createCell(++colIdx);
			cell.setCellValue(arrs);
			
		}
		// �q getData() �^�Ǫ��r��}�C�A�b Excel ��(B,2)�}�l�̧ǫ����g�J
		String[] strData = getData();
		for (String arrs : strData) {
			row = sheet.createRow(++rowIdx);
			cell = row.createCell(1);
			cell.setCellValue(arrs);
			sheet.autoSizeColumn(rowIdx); // �۰ʽվ����e��
		}
		
		// ���w�ɮצW��
		String fileName = "test_ws_data.xlsx";

		/*
		 * �|�����w�ɮ׸��|�A�ɮ׫إߦb������M�פ� �x�s�u�@ï
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

	// �ϥ� Jsoup �����ӫ~�W��
	public static String[] getData() {
		String strUrl = "https://www.woodstuck.com.tw/categories/%E5%A4%96%E5%A5%97-1";
		String str = null;
		try {
			Document doc = Jsoup.connect(strUrl).timeout(30000).validateTLSCertificates(false).get();

			Elements el1 = doc.select("ul.boxify-container");
			// ## ���o�ӭ����Ҧ��ӫ~�W��
			for (Element postItem : el1) {
				Elements el = postItem.select("div.title");
				str = el.text();
				System.out.println(el.text());
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
		// str �����e�O"�ӫ~�W��"+" "���@���s��ʪ��r��A�]���n�����ΡA�é�i String[] data ��
		String[] data = str.split(" ");
		return data;
	}

}
