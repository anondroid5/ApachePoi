package sample.poi.read;


import java.io.*;

//Apache POI ライブラリ群を全選択
import org.apache.poi.ss.usermodel.*;

/*apache poiを使用した基本的な読み込み用サンプルプログラム*/
public class Read {
	public static void main(String[] args) throws Exception {
	    // Excelファイルの指定
	    Workbook wb = WorkbookFactory.create(new FileInputStream("./files/read.xlsx"));

	  //エクセルのシートの情報を表示する
	    System.out.println("総シート数:" + wb.getNumberOfSheets());
	    System.out.println("シート名(1):" + wb.getSheetName(0));
		System.out.println("シート名(2):" + wb.getSheetName(1));
		System.out.println("シート名(3):" + wb.getSheetName(2));

		//引数で指定した番号のシートを取得する
	    Sheet sheet = wb.getSheetAt(0);
	    
	    Row row = sheet.getRow(0);//左から1番目のセル
	    Cell cell = row.getCell(0); //(0,0)を読み込む
	    System.out.print("左上セル(0,0)の内容＝");
	    System.out.println(cell.getStringCellValue());
	  }

}
