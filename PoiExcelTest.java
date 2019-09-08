package poiexceltest;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class PoiExcelTest {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		String file=args[0];
		String file1=args[1];
		Sheet sheet1 =CreateWb.createsheet(file);
		Sheet sheet2 =CreateWb.createsheet(file1);
		System.out.println(sheet1.getSheetName());
		System.out.println(sheet2.getSheetName());
		String[] financial= {
				"应收帐款","存货","流动资产合计","资产总计","流动负债合计","负债合计","所有者权益合计","一、商品销售收入","二、商品销售利润","三、营业利润","五、净利润","负债及所有者权益总计","四、利润总额"};
		double[] finanvalue1=CreateWb.handlenum(sheet1, financial, file);
		double[] finanvalue2=CreateWb.handlenum(sheet2, financial, file1);
		double[] finanvalue=CreateWb.MergeArrays(finanvalue1,finanvalue2);
		double[] numc=CreateWb.Algorithm(finanvalue);
		CreateWb.printfile(numc, args[2],finanvalue1,finanvalue2);
		//已经打开excel表格
		
		
	}

}
