package sample.poi.write;

import java.io.*;

//Apache POI ライブラリ群
import org.apache.poi.ss.usermodel.*;

/*read.xlsxをwrite.xlsxに書き込むサンプルプログラム*/
public class Write {
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
	    
	    Row row = sheet.getRow(0);
	    Cell cell = row.getCell(0); // (0,0)
	    System.out.print("左上セル(0,0)の内容＝");
	    System.out.println(cell.getStringCellValue());
	    
	    // 上記セルに値を設定
	    cell.setCellValue("bbbbb");
	    
	    // 他のセルに値を設定(無いところに作る場合は作成してから）
	    cell = row.createCell(1); // (1,0)
	    cell.setCellValue(10);
	    cell = row.createCell(2); // (2,0)
	    cell.setCellValue(1.25);
	    cell = row.createCell(3); // (3,0)
	    cell.setCellValue("文字列書き込み");
	    cell = row.createCell(4); // (4,0)
	    cell.setCellValue("Test");
	    
	    // 罫線の設定。上下左右バラバラに罫線を指定できる。
	    // 罫線はスタイルとして指定しておき、あとからセルに設定する
	    CellStyle style = wb.createCellStyle();
	    // セルの左と右の線。線の種類を指定している
	    style.setBorderLeft(CellStyle.BORDER_DASHED);
	    style.setBorderRight(CellStyle.BORDER_DOUBLE);
	    // セル上部の線。色を指定している
	    style.setBorderBottom(CellStyle.BORDER_MEDIUM);
	    style.setBottomBorderColor(IndexedColors.MAROON.getIndex());
	    
	    // セルの背景色
	    // 背景色はスタイルとして指定しておき、あとからセルに設定する
	    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    style.setFillForegroundColor(IndexedColors.SKY_BLUE.getIndex());
	    
	    // セルに対してスタイルを設定
	    cell.setCellStyle(style);
	    
	    // 書き込み
	    FileOutputStream out = new FileOutputStream("./files/write.xlsx");
	    wb.write(out);
	  }

}
