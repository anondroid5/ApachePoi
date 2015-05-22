package sample.poi.read;

import java.io.*;

//Apache POI ライブラリ群
import org.apache.poi.ss.usermodel.*;

public class SoccerPlayerReader {
	public static void main(String[] args) throws Exception {
	    Workbook wb = WorkbookFactory.create(new FileInputStream("./files/soccerplayer.xlsx"));
	    
	    //エクセルのシートの情報を表示する
	    System.out.println("総シート数:" + wb.getNumberOfSheets());
	    System.out.println("シート名(1):" + wb.getSheetName(0));
		System.out.println("シート名(2):" + wb.getSheetName(1));
		System.out.println("シート名(3):" + wb.getSheetName(2));

		//引数で指定した番号のシートを取得する
	    Sheet sheet = wb.getSheetAt(0);
	    
	    for(int i=0; i<45; i++){
	    	Row row = sheet.getRow(i);//i行目を読み込む
	    	Cell cell = row.getCell(0); // (0,i)となる。i行目の0番目のセルを読み込む
	    	System.out.print("セル(0,"+i+")のサッカー選手名:");
	    	System.out.println(cell.getStringCellValue());
	    }
	 }
}
