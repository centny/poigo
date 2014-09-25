package poigo;

import org.apache.poi.ss.usermodel.CellValue;

public class Show {

	public static void show(String s) {
		System.out.println(s);
	}
	public static void show(CellValue s) {
		System.out.println(s.formatAsString());
	}
	public static void main(String[] args) {

	}
}
