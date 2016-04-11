package poi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

/**
 * @Description:
 * @Author:	nutony
 * @Company:	http://java.itcast.cn
 * @CreateDate:	2014年10月14日
 */
public class test {

	@Test
	public void testHSSF_base() throws IOException{
		/*
		 * 开发步骤：
		 * 1、创建一个工作簿
		 * 2、创建一个工作表
		 * 3、创建一个行对象
		 * 4、创建一个单元格对象，指定它的列
		 * 5、给单元格设置内容
		 * 6、样式进行修饰（跳过）
		 * 7、保存，写文件
		 * 8、关闭对象
		 */
		
		Workbook wb = new HSSFWorkbook();
		Sheet sheet = wb.createSheet();
		Row nRow = sheet.createRow(7);			//第八行
		Cell nCell = nRow.createCell(4);		//第五列
		
		nCell.setCellValue("传智播客万年长！");
		
		OutputStream os = new FileOutputStream("c:\\testpoi.xls");	//excel 2003
		wb.write(os);
		
		os.flush();
		os.close();
	}
	
	@Test	//带样式
	public void testHSSFStyle() throws IOException{
		Workbook wb = new HSSFWorkbook();
		Sheet sheet = wb.createSheet();
		Row nRow = sheet.createRow(7);			//第八行
		Cell nCell = nRow.createCell(4);		//第五列
		
		nCell.setCellValue("传智播客万年长！");
		
		//设置样式
		CellStyle titleStyle = wb.createCellStyle();		//创建单元格样式
		Font titleFont = wb.createFont();					//创建一个字体对象
		
		titleFont.setFontName("华文隶书");					//设置字体名称
		titleFont.setFontHeightInPoints((short)24);			//设置字体大小
		
		titleStyle.setFont(titleFont);						//绑定字体对象
		
		
		nCell.setCellStyle(titleStyle);						//设置单元格样式
		
		Row xRow = sheet.createRow(8);
		Cell xCell = xRow.createCell(6);
		
		CellStyle textSytle = wb.createCellStyle();
		Font textFont = wb.createFont();
		
		textFont.setFontName("Times New Roman");
		textFont.setFontHeightInPoints((short)14);
		
		textSytle.setFont(textFont);
		
		xCell.setCellValue("java.itcast.cn");
		xCell.setCellStyle(textSytle);
		
		OutputStream os = new FileOutputStream("c:\\testpoi.xls");	//excel 2003
		wb.write(os);
		
		os.flush();
		os.close();
	}
}
