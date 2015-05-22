package sample.poi.read;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Dictionary {
	public static void main(String[] args) throws Exception {
		//読み込みファイルを指定
	    Workbook wb = WorkbookFactory.create(new FileInputStream("./files/dict.xlsx"));
	    
	    //エクセルのシートの情報を表示する
	    System.out.println("総シート数:" + wb.getNumberOfSheets());
	    System.out.println("シート名(1):" + wb.getSheetName(0));
		System.out.println("シート名(2):" + wb.getSheetName(1));
		System.out.println("シート名(3):" + wb.getSheetName(2));
		System.out.println();//空白用

		//0番のシートを取得する
	    Sheet sheet0 = wb.getSheetAt(0);
	    
	    System.out.println(wb.getSheetName(0)+"のシートを可視化");
	    for(int i=0; i<106; i++){
	    	Row row = sheet0.getRow(i);//i行目を読み込む
	    	Cell cell = row.getCell(0); // (0,i)となる。i行目の0番目のセルを読み込む
	    	System.out.print("セル(0,"+i+")のサッカー専門用語:");
	    	System.out.println(cell.getStringCellValue());
	    }
	    System.out.println();//空白用
	    //1番のシートを取得する
	    Sheet sheet1 = wb.getSheetAt(1);
	    
	    System.out.println(wb.getSheetName(1)+"のシートを可視化");
	    for(int i=0; i<23; i++){
	    	Row row = sheet1.getRow(i);//i行目を読み込む
	    	Cell cell0 = row.getCell(0); // (0,i)となる。i行目の0番目のセルを読み込む
	    	Cell cell1 = row.getCell(1); // (1,i)となる。i行目の0番目のセルを読み込む
	    	System.out.print("セル(0,"+i+")のサッカー選手名:");
	    	System.out.println(cell0.getStringCellValue()+cell1.getStringCellValue());
	    }
	    
	    
	 }
}
