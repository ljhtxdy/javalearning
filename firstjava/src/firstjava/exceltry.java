package firstjava;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Date;
import java.util.Iterator;
import javax.swing.filechooser.FileSystemView;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.charts.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.charts.*;

public class exceltry {
	
	public static void main(String[] args) throws Exception {
		//hssfcreateExcel();
		//hssfreadExcel();
		xssfcreate();
		//readbyrow();
		plotchart();
	}
	
	public static void hssfreadExcel() throws IOException{
	    FileSystemView fsv = FileSystemView.getFileSystemView();
	    String desktop = fsv.getHomeDirectory().getPath();
	    String filePath = desktop + "/template.xls";

	    FileInputStream fileInputStream = new FileInputStream(filePath);
	    BufferedInputStream bufferedInputStream = new BufferedInputStream(fileInputStream);
	    POIFSFileSystem fileSystem = new POIFSFileSystem(bufferedInputStream);
	    HSSFWorkbook workbook = new HSSFWorkbook(fileSystem);
	    HSSFSheet sheet = workbook.getSheet("Sheet1");

	    int lastRowIndex = sheet.getLastRowNum();
	    System.out.println(lastRowIndex);
	    for (int i = 0; i <= lastRowIndex; i++) {
	        HSSFRow row = sheet.getRow(i);
	        if (row == null) { break; }

	        short lastCellNum = row.getLastCellNum();
	        for (int j = 0; j < lastCellNum; j++) {
	            Object cellValue = row.getCell(j).getStringCellValue();
	            System.out.println(cellValue);
	        }
	    }


	    bufferedInputStream.close();
	}

	public static void hssfcreateExcel() throws IOException{
	    // 获取桌面路径
	    FileSystemView fsv = FileSystemView.getFileSystemView();
	    String desktop = fsv.getHomeDirectory().getPath();
	    String filePath = desktop + "/template.xls";

	    File file = new File(filePath);
	    OutputStream outputStream = new FileOutputStream(file);
	    HSSFWorkbook workbook = new HSSFWorkbook();
	    HSSFSheet sheet = workbook.createSheet("Sheet1");
	    HSSFRow row = sheet.createRow(0);
	    row.createCell(0).setCellValue("id");
	    row.createCell(1).setCellValue("订单号");
	    row.createCell(2).setCellValue("下单时间");
	    row.createCell(3).setCellValue("个数");
	    row.createCell(4).setCellValue("单价");
	    row.createCell(5).setCellValue("订单金额");
	    row.setHeightInPoints(30); // 设置行的高度

	    HSSFRow row1 = sheet.createRow(1);
	    row1.createCell(0).setCellValue("1");
	    row1.createCell(1).setCellValue("NO00001");

	    // 日期格式化
	    HSSFCellStyle cellStyle2 = workbook.createCellStyle();
	    HSSFCreationHelper creationHelper = workbook.getCreationHelper();
	    cellStyle2.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
	    sheet.setColumnWidth(2, 20 * 256); // 设置列的宽度

	    HSSFCell cell2 = row1.createCell(2);
	    cell2.setCellStyle(cellStyle2);
	    cell2.setCellValue(new Date());

	    row1.createCell(3).setCellValue(2);


	    // 保留两位小数
	    HSSFCellStyle cellStyle3 = workbook.createCellStyle();
	    cellStyle3.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
	    HSSFCell cell4 = row1.createCell(4);
	    cell4.setCellStyle(cellStyle3);
	    cell4.setCellValue(29.5);


	    // 货币格式化
	    HSSFCellStyle cellStyle4 = workbook.createCellStyle();
	    HSSFFont font = workbook.createFont();
	    font.setFontName("华文行楷");
	    font.setFontHeightInPoints((short)15);
	    font.setColor(HSSFColor.RED.index);
	    cellStyle4.setFont(font);

	    HSSFCell cell5 = row1.createCell(5);
	    cell5.setCellFormula("D2*E2");  // 设置计算公式

	    // 获取计算公式的值
	    HSSFFormulaEvaluator e = new HSSFFormulaEvaluator(workbook);
	    cell5 = e.evaluateInCell(cell5);
	    System.out.println(cell5.getNumericCellValue());


	    workbook.setActiveSheet(0);
	    workbook.write(outputStream);
	    outputStream.close();
	}
	
	public static void xssfopen() throws Exception
	   { 
	      File file = new File("C:\\Users\\ljh\\Desktop\\template.xlsx");
	      
	      FileInputStream fIP = new FileInputStream(file);
	      //Get the workbook instance for XLSX file 
	      XSSFWorkbook workbook = new XSSFWorkbook(fIP);
	      if(file.isFile() && file.exists())
	      {
	         System.out.println(
	         "openworkbook.xlsx file open successfully.");
	      }
	      else
	      {
	         System.out.println(
	         "Error to open workbook.xlsx file.");
	      }
	   }
	
	public static void xssfcreate()throws Exception 
	   {
	      //Create Blank workbook
	      XSSFWorkbook workbook = new XSSFWorkbook(); 
	      
	      XSSFSheet spreadsheet=workbook.createSheet("sb");
	      
	    //create first row on a created spreadsheet
	      XSSFRow row = spreadsheet.createRow(0);
	      //create first cell on created row
	      XSSFCell cell = row.createCell(0);
	      
	      row = spreadsheet.createRow((short) 2);
	      row.createCell(0).setCellValue("Type of Cell");
	      row.createCell(1).setCellValue("cell value");
	      row = spreadsheet.createRow((short) 3);
	      row.createCell(0).setCellValue("set cell type BLANK");
	      row.createCell(1);
	      row = spreadsheet.createRow((short) 4);
	      row.createCell(0).setCellValue("set cell type BOOLEAN");
	      row.createCell(1).setCellValue(true);
	      row = spreadsheet.createRow((short) 5);
	      row.createCell(0).setCellValue("set cell type ERROR");
	      row.createCell(1).setCellValue(XSSFCell.CELL_TYPE_ERROR );
	      row = spreadsheet.createRow((short) 6);
	      row.createCell(0).setCellValue("set cell type date");
	      row.createCell(1).setCellValue(new Date());
	      row = spreadsheet.createRow((short) 7);
	      row.createCell(0).setCellValue("set cell type numeric" );
	      row.createCell(1).setCellValue(20 );
	      row = spreadsheet.createRow((short) 8);
	      row.createCell(0).setCellValue("set cell type string");
	      row.createCell(1).setCellValue("A String");
	      
	      //Create file system using specific name
	      FileOutputStream out = new FileOutputStream(
	      new File("D:\\createworkbook.xlsx"));
	      //write operation workbook using file out object 
	      workbook.write(out);
	      out.close();
	      System.out.println("createworkbook.xlsx written successfully");
	   }
	
	static XSSFRow row;
	
	public static void readbyrow() throws Exception 
	   {
	      FileInputStream fis = new FileInputStream(new File("C:\\Users\\ljh\\Desktop\\template.xlsx"));
	      XSSFWorkbook workbook = new XSSFWorkbook(fis);
	      XSSFSheet spreadsheet = workbook.getSheetAt(0);
	      Iterator < Row > rowIterator = spreadsheet.iterator();
	      while (rowIterator.hasNext()) 
	      {
	         row = (XSSFRow) rowIterator.next();
	         Iterator < Cell > cellIterator = row.cellIterator();
	         while ( cellIterator.hasNext()) 
	         {
	            Cell cell = cellIterator.next();
	            switch (cell.getCellType()) 
	            {
	               case Cell.CELL_TYPE_NUMERIC:
	               System.out.print( 
	               cell.getNumericCellValue() + " \t\t " );
	               break;
	               case Cell.CELL_TYPE_STRING:
	               System.out.print(
	               cell.getStringCellValue() + " \t\t " );
	               break;
	            }
	         }
	         System.out.println();
	      }
	      fis.close();
	   }
	
	 public static void plotchart()throws Exception
		   {
		                         /* Create a Workbook object that will hold the final chart */
		                        XSSFWorkbook my_workbook = new XSSFWorkbook();
		                         /* Create a worksheet object for the line chart. This worksheet will contain the chart */
		                        XSSFSheet my_worksheet = my_workbook.createSheet("LineChart_Example");
		                        
		                        /* Let us now create some test data for the chart */
		                        /* Later we can see how to get this test data from a CSV File or SQL Table */
		                        /* We use a 4 Row chart input with 5 columns each */
		                        for (int rowIndex = 0; rowIndex < 4; rowIndex++)
		                {
		                        /* Add a row that contains the chart data */
		                        XSSFRow my_row = my_worksheet.createRow((short)rowIndex);
		                        for (int colIndex = 0; colIndex < 5; colIndex++)
		                {
		                        /* Define column values for the row that is created */
		                        XSSFCell cell = my_row.createCell((short)colIndex);
		                        cell.setCellValue(colIndex * (rowIndex + 1));
		                }
		                }               
		                        /* At the end of this step, we have a worksheet with test data, that we want to write into a chart */
		                        /* Create a drawing canvas on the worksheet */
		                        XSSFDrawing xlsx_drawing = my_worksheet.createDrawingPatriarch();
		                        /* Define anchor points in the worksheet to position the chart */
		                        XSSFClientAnchor anchor = xlsx_drawing.createAnchor(0, 0, 0, 0, 0, 5, 10, 15);
		                        /* Create the chart object based on the anchor point */
		                        XSSFChart my_line_chart = xlsx_drawing.createChart(anchor);
		                        /* Define legends for the line chart and set the position of the legend */
		                        XSSFChartLegend legend = my_line_chart.getOrCreateLegend();
		                        legend.setPosition(LegendPosition.BOTTOM);     
		                        /* Create data for the chart */
		                        LineChartData data = my_line_chart.getChartDataFactory().createLineChartData();     
		                        /* Define chart AXIS */
		                        ChartAxis bottomAxis = my_line_chart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
		                        ValueAxis leftAxis = my_line_chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
		                        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);     
		                        /* Define Data sources for the chart */
		                        /* Set the right cell range that contain values for the chart */
		                        /* Pass the worksheet and cell range address as inputs */
		                        /* Cell Range Address is defined as First row, last row, first column, last column */
		                        ChartDataSource<Number> xs = DataSources.fromNumericCellRange(my_worksheet, new CellRangeAddress(0, 0, 0, 4));
		                        ChartDataSource<Number> ys1 = DataSources.fromNumericCellRange(my_worksheet, new CellRangeAddress(1, 1, 0, 4));
		                        ChartDataSource<Number> ys2 = DataSources.fromNumericCellRange(my_worksheet, new CellRangeAddress(2, 2, 0, 4));
		                        ChartDataSource<Number> ys3 = DataSources.fromNumericCellRange(my_worksheet, new CellRangeAddress(3, 3, 0, 4));
		                        /* Add chart data sources as data to the chart */
		                        data.addSeries(xs, ys1);
		                        data.addSeries(xs, ys2);
		                        data.addSeries(xs, ys3);
		                        /* Plot the chart with the inputs from data and chart axis */
		                        my_line_chart.plot(data, new ChartAxis[] { bottomAxis, leftAxis });
		                        /* Finally define FileOutputStream and write chart information */               
		                        FileOutputStream fileOut = new FileOutputStream("D:\\xlsx-line-chart.xlsx");
		                        my_workbook.write(fileOut);
		                        fileOut.close();
		   }
	
	
	
	
	

	  
}
		
		
	


