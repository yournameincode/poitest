package poiexceltest;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class CreateWb {
	public static Sheet createsheet(String path) throws IOException {
		InputStream File=new FileInputStream(path);
		Workbook wb=new HSSFWorkbook(File);
		Sheet sheet=wb.getSheetAt(0);
		File.close();
		return sheet;
	}
	public static double[] MergeArrays(double[] a,double[] b) {
		double[] finanvalue=new double[16];
		for(int i=0;i<16;i++) {
			if(i<11) {finanvalue[i]=a[i];}
			else {
				finanvalue[i]=b[i-11];
			}
			
		}
		return finanvalue;
	}
	public static double[] Algorithm(double[] a) {

		double[] numC=new double [9];
		numC[0]=a[8]/a[5]*100;
		numC[1]=(a[3]-a[2])/a[7]*100;
		numC[2]=a[15]/a[5];
		numC[3]=a[14]/(a[0]+a[1])/2;
		numC[4]=a[15]/(a[3]+a[4])/2;
		numC[5]=a[15]/(a[5]+a[6])/2*100;
		
		numC[6]=a[14]/a[13]*100;
		numC[7]=(a[11]-a[12])/a[11]*100;
		numC[8]=a[10]/a[9]*100;
		return numC;
	}
	public static void printfile(double[] numC,String path ,double[] finanvalue1,double[] finanvalue2) throws IOException {
		PrintWriter out=new PrintWriter(new FileOutputStream(path),true);
		out.println(String.format("%.2f", numC[0])+"%");
		out.println(String.format("%.2f", numC[1])+"%");
		out.println(String.format("%.2f", numC[2]));
		out.println(String.format("%.2f", numC[3]));
		out.println(String.format("%.2f", numC[4]));
		out.println(String.format("%.2f", numC[5])+"%");
		out.println(String.format("%.2f", numC[6])+"%");
		out.println(String.format("%.2f", numC[5])+"%");
		out.println(String.format("%.2f", numC[7])+"%");
		out.println(String.format("%.2f", numC[8])+"%");
		out.println(String.format("%.2f", numC[7])+"%");
		out.println();
		out.println();
		out.println(String.format("%.2f", numC[7])+"%");
		out.println(String.format("%.2f", numC[8])+"%");
		out.println(String.format("%.2f", numC[7])+"%");
		out.println();
		out.println();
		out.println(String.format("%.2f", numC[0])+"%");
		out.println(String.format("%.2f", numC[1])+"%");
		out.println();
		out.println();
		out.println(String.format("%.2f", numC[2]));
		out.println(String.format("%.2f", numC[3]));
		out.println(String.format("%.2f", numC[4]));
		out.println();
		out.println();
		out.println(String.format("%.2f", numC[5])+"%");
		out.println(String.format("%.2f", numC[6])+"%");
		out.println(String.format("%.2f", numC[5])+"%");
		out.println();
		out.println("资产负债表");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println(String.format("%.2f", finanvalue1[0]));
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println(String.format("%.2f", finanvalue1[2]));//9
		out.println("--");
		out.println("--");
		out.println(String.format("%.2f", finanvalue1[3]));//12
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");//26
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("资产负债表的第二页");
		out.println(String.format("%.2f", finanvalue1[5]));//37
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println(String.format("%.2f", finanvalue1[7]));//50
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println(String.format("%.2f", finanvalue1[8]));//60
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println(String.format("%.2f", finanvalue1[9]));;//67
		out.println(String.format("%.2f", finanvalue1[11]));//68
		out.println();
		out.println("利润表");
		out.println(String.format("%.2f", finanvalue2[0]));//1
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");
		out.println("--");//10
		out.println("--");
		out.println("--");
		out.println("--");
		out.println(String.format("%.2f", finanvalue2[3]));//14
		out.println("--");
		out.println("--");
		out.println("--");//17
		out.println(String.format("%.2f", finanvalue2[5]));
		out.println("--");
		out.println(String.format("%.2f", finanvalue2[4]));
		out.println("--");
		out.println("--");
		out.println("--");
		out.close();
	}
	public static double[] handlenum(Sheet sheet,String[] financial,String path) {
		int rownum=sheet.getLastRowNum();
		double[] finanvalue=new double[12];
		for(int startrownum=0;startrownum<rownum;startrownum++) {
			Row row=sheet.getRow(startrownum);
			if(path.indexOf("资产负债表")>-1) {
				String cell=row.getCell(0).getRichStringCellValue().toString().trim();
				String cell1=row.getCell(4).getRichStringCellValue().toString().trim();
				if(cell.equals(financial[0])) {
					//在java中比较两个字符串是否相等要使用equals
					finanvalue[0]=row.getCell(2).getNumericCellValue();
					finanvalue[1]=row.getCell(3).getNumericCellValue();
				}
				if(cell.equals(financial[1])) {
					//在java中比较两个字符串是否相等要使用equals
					finanvalue[2]=row.getCell(2).getNumericCellValue();
				}
				if(cell.equals(financial[2])) {
					//在java中比较两个字符串是否相等要使用equals
					finanvalue[3]=row.getCell(2).getNumericCellValue();
					finanvalue[4]=row.getCell(3).getNumericCellValue();
				}
				if(cell.equals(financial[3])) {
					//在java中比较两个字符串是否相等要使用equals
					finanvalue[5]=row.getCell(2).getNumericCellValue();
					finanvalue[6]=row.getCell(3).getNumericCellValue();
				}
				if(cell1.equals(financial[4])) {
					//在java中比较两个字符串是否相等要使用equals
					finanvalue[7]=row.getCell(6).getNumericCellValue();
				}
				if(cell1.equals(financial[5])) {
					//在java中比较两个字符串是否相等要使用equals
					finanvalue[8]=row.getCell(6).getNumericCellValue();
				}
				if(cell1.equals(financial[6])) {
					//在java中比较两个字符串是否相等要使用equals
					finanvalue[9]=row.getCell(6).getNumericCellValue();
					finanvalue[10]=row.getCell(7).getNumericCellValue();
				}
				if(cell1.equals(financial[11])) {
					//在java中比较两个字符串是否相等要使用equals
					finanvalue[11]=row.getCell(6).getNumericCellValue();		
				}
			}
			else {
				String cell=row.getCell(0).getRichStringCellValue().toString().trim();
				if(cell.equals(financial[7])) {
					//在java中比较两个字符串是否相等要使用equals
					finanvalue[0]=row.getCell(2).getNumericCellValue();
					finanvalue[1]=row.getCell(3).getNumericCellValue();
				}
				if(cell.equals(financial[8])) {
					//在java中比较两个字符串是否相等要使用equals
					finanvalue[2]=row.getCell(2).getNumericCellValue();
				}
				if(cell.equals(financial[9])) {
					//在java中比较两个字符串是否相等要使用equals
					finanvalue[3]=row.getCell(2).getNumericCellValue();
					
				}
				if(cell.equals(financial[10])) {
					//在java中比较两个字符串是否相等要使用equals
					finanvalue[4]=row.getCell(2).getNumericCellValue();	
				}
				if(cell.equals(financial[12])) {
					//在java中比较两个字符串是否相等要使用equals
					finanvalue[5]=row.getCell(2).getNumericCellValue();	
				}
			}
		
		//System.out.println(rownum);//成功显示第一个sheet表格的行数		
			//System.out.println(cell);
			//System.out.println(cell1.getRichStringCellValue().toString().trim());
			//已经获取到了每一行第一个单元格的数据

		}
		return finanvalue;
	}
}
