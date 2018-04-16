package firstjava;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import javax.swing.filechooser.FileSystemView;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.charts.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.charts.*;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class exceltry {

	public static void main(String[] args) throws Exception {
		// hssfcreateExcel();
		// hssfreadExcel();
		// xssfcreate();
		// readbyrow();
		// plotchart();
		complete();
		// test();

	}

	public static void hssfreadExcel() throws IOException {
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
			if (row == null) {
				break;
			}

			short lastCellNum = row.getLastCellNum();
			for (int j = 0; j < lastCellNum; j++) {
				Object cellValue = row.getCell(j).getStringCellValue();
				System.out.println(cellValue);
			}
		}

		bufferedInputStream.close();
	}

	public static void hssfcreateExcel() throws IOException {
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
		font.setFontHeightInPoints((short) 15);
		font.setColor(HSSFColor.RED.index);
		cellStyle4.setFont(font);

		HSSFCell cell5 = row1.createCell(5);
		cell5.setCellFormula("D2*E2"); // 设置计算公式

		// 获取计算公式的值
		HSSFFormulaEvaluator e = new HSSFFormulaEvaluator(workbook);
		cell5 = e.evaluateInCell(cell5);
		System.out.println(cell5.getNumericCellValue());

		workbook.setActiveSheet(0);
		workbook.write(outputStream);
		outputStream.close();
	}

	// only used to open a excel file

	public static void xssfopen() throws Exception {
		File file = new File("C:\\Users\\ljh\\Desktop\\template.xlsx");

		FileInputStream fIP = new FileInputStream(file);
		// Get the workbook instance for XLSX file
		XSSFWorkbook workbook = new XSSFWorkbook(fIP);
		if (file.isFile() && file.exists()) {
			System.out.println("openworkbook.xlsx file open successfully.");
		} else {
			System.out.println("Error to open workbook.xlsx file.");
		}
	}

	// used to create a new file and then write the data
	public static void xssfcreate() throws Exception {
		// Create Blank workbook
		XSSFWorkbook workbook = new XSSFWorkbook();

		XSSFSheet spreadsheet = workbook.createSheet("sb");

		// create first row on a created spreadsheet
		XSSFRow row = spreadsheet.createRow(0);
		// create first cell on created row
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
		row.createCell(1).setCellValue(XSSFCell.CELL_TYPE_ERROR);
		row = spreadsheet.createRow((short) 6);
		row.createCell(0).setCellValue("set cell type date");
		row.createCell(1).setCellValue(new Date());
		row = spreadsheet.createRow((short) 7);
		row.createCell(0).setCellValue("set cell type numeric");
		row.createCell(1).setCellValue(20);
		row = spreadsheet.createRow((short) 8);
		row.createCell(0).setCellValue("set cell type string");
		row.createCell(1).setCellValue("A String");

		// Create file system using specific name
		FileOutputStream out = new FileOutputStream(new File("D:\\createworkbook.xlsx"));
		// write operation workbook using file out object
		workbook.write(out);
		out.close();
		System.out.println("createworkbook.xlsx written successfully");
	}

	static XSSFRow row;

	public static void readbyrow() throws Exception {
		FileInputStream fis = new FileInputStream(new File("C:\\Users\\ljh\\Desktop\\template.xlsx"));
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet spreadsheet = workbook.getSheetAt(0);
		Iterator<Row> rowIterator = spreadsheet.iterator();
		while (rowIterator.hasNext()) {
			row = (XSSFRow) rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_NUMERIC:
					System.out.print(cell.getNumericCellValue() + " \t\t ");
					break;
				case Cell.CELL_TYPE_STRING:
					System.out.print(cell.getStringCellValue() + " \t\t ");
					break;
				}
			}
			System.out.println();
		}
		fis.close();
	}

	public static void plotchart() throws Exception {
		/* Create a Workbook object that will hold the final chart */
		XSSFWorkbook my_workbook = new XSSFWorkbook();
		/*
		 * Create a worksheet object for the line chart. This worksheet will contain the
		 * chart
		 */
		XSSFSheet my_worksheet = my_workbook.createSheet("LineChart_Example");

		/* Let us now create some test data for the chart */
		/* Later we can see how to get this test data from a CSV File or SQL Table */
		/* We use a 4 Row chart input with 5 columns each */
		for (int rowIndex = 0; rowIndex < 4; rowIndex++) {
			/* Add a row that contains the chart data */
			XSSFRow my_row = my_worksheet.createRow((short) rowIndex);
			for (int colIndex = 0; colIndex < 5; colIndex++) {
				/* Define column values for the row that is created */
				XSSFCell cell = my_row.createCell((short) colIndex);
				cell.setCellValue(colIndex * (rowIndex + 1));
			}
		}
		/*
		 * At the end of this step, we have a worksheet with test data, that we want to
		 * write into a chart
		 */
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
		/*
		 * Cell Range Address is defined as First row, last row, first column, last
		 * column
		 */
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

	public static void complete() throws Exception {

		// read the file

		File file = new File("C:\\Users\\ljh\\Desktop\\all.xlsx");

		FileInputStream fIP = new FileInputStream(file);
		// Get the workbook instance for XLSX file
		XSSFWorkbook workbook = new XSSFWorkbook(fIP);
		if (file.isFile() && file.exists()) {
			System.out.println("openworkbook.xlsx file open successfully.");
		} else {
			System.out.println("Error to open workbook.xlsx file.");
		}

		XSSFSheet sheet = workbook.getSheetAt(0); // 获取第一个工作表的对象
		XSSFSheet plotsheet = workbook.createSheet("LineChart");

		// delete the front part row 1-14
		for (int i = 0; i <= 13; i++) {
			sheet.shiftRows(1, sheet.getLastRowNum(), -1);
		}

		// delete the back part
		int index = 0;

		for (int i = 0; i <= sheet.getLastRowNum(); i++) {
			XSSFRow row = sheet.getRow(i); // 获取每一行的对象

			if (row == null) {
				index = i;
				break;
			}
		}

		int last;

		last = sheet.getLastRowNum();

		int num;

		num = last - index;

		for (int i = 0; i <= num; i++) {
			sheet.shiftRows(index + 1, last, -1);
		}

		// adjust wrong format data
		int index1 = 0;
		int index2=0;
		ArrayList<Integer> part = new ArrayList<Integer>();
		int [][] sequence = new int[100][];

		for (int i = sheet.getLastRowNum(); i >= 2; i--) {
			XSSFRow row = sheet.getRow(i); // 获取每一行的对象
			int number = 0;
			ArrayList<Integer> List = new ArrayList<Integer>();
			int decide = 0;

			for (int j = 0; j <= row.getLastCellNum(); j++) {

				XSSFCell cell = row.getCell(j); // 获取每一行的每一列
				if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK) {
					number = number + 1;
					List.add(j);
				}
			}

			XSSFRow lastrow = sheet.getRow(i - 1); // 获取每一行的对象
			
			for (int k = 0; k < number; k++) {
				XSSFCell lastcell = lastrow.getCell(List.get(k));
				if (lastcell == null || lastcell.getCellType() == Cell.CELL_TYPE_BLANK) {
					decide = 1;
					sequence[index2]=new int[number];

					for (int l = 0; l <= number - 1; l++) {

						sequence[index2][l] =List.get(l);
					}
					
					break;

				}

				
			}

			if (decide == 1) {
				index1 = index1 + 1;
				index2 = index2 + 1;
				part.add(i);
			}

		}
		
		
		XSSFSheet newsheet = workbook.cloneSheet(0, "raw data"); // COPY第一个工作表的对象

		for (int i = part.get(index1-1); i <= newsheet.getLastRowNum(); i++) {
			XSSFRow row = newsheet.getRow(i);
			if (row != null) {
			newsheet.removeRow(row);
			}
		}

		for(int i=0;i<part.size();i++) {
			System.out.println(part.get(i));
		}
		
		
		
		
		
		
		
		
		
		
		
		
		/*for (int i = index1-1; i >= 0; i--) {
			XSSFRow startrow = sheet.getRow(1);
			XSSFRow startrow1 = newsheet.getRow(1);
			XSSFRow finalrow = sheet.getRow(1);
			XSSFRow finalrow1 = newsheet.getRow(1);

			startrow = sheet.getRow(part.get(i));
			startrow1 = newsheet.getRow(part.get(i));

			int start = 0;
			int final_ = 0;
			

			start = part.get(i);

			if (i == 0) {
				finalrow = sheet.getRow(sheet.getLastRowNum());
				finalrow1 = newsheet.getRow(sheet.getLastRowNum());

				final_ = sheet.getLastRowNum();

			} else {

				finalrow = sheet.getRow(part.get(i-1) - 1);
				finalrow1 = newsheet.getRow(part.get(i - 1) - 1);

				final_ = part.get(i - 1) - 1;

			}

			ArrayList<Integer> List_ = new ArrayList<Integer>();

			for(int t=0;t<sequence[i].length;t++)
			{
			List_.add (sequence[i][t]);
			}

			XSSFCell firstcell = startrow.getCell(List_.get(0));

			Double data = firstcell.getNumericCellValue();

			XSSFRow prerow = newsheet.getRow(part.get(i) - 1);

			ArrayList<Integer> prelist = new ArrayList<Integer>();
			
			int number123 = 0;

			for (int f = 0; f < prerow.getLastCellNum(); f++) {

				XSSFCell precell = prerow.getCell(f);
				// 获取每一行的每一列
				if (precell != null && precell.getCellType() != Cell.CELL_TYPE_BLANK) {
					number123 = number123 + 1;
					prelist.add(f);
				}
			}
			
			int preindex = 0;
			
			for (int m = 0; m < number123; m++) {
				
				Double predata = prerow.getCell(prelist.get(m)).getNumericCellValue();

				if ((data - predata) <= 2 && (data - predata)>0) {
					preindex = m;
					break;

				}

			}

			int delta = List_.get(0) - prelist.get(preindex);

			for (int b = start; b <= final_; b++) {
				for (int c = 0; c <= List_.size(); c++) {

					if (sheet.getRow(b).getCell(c) != null
							&& sheet.getRow(b).getCell(c).getCellType() != Cell.CELL_TYPE_BLANK) {
						Double value = sheet.getRow(b).getCell(c).getNumericCellValue();
						newsheet.getRow(b).getCell(c - delta).setCellValue(value);
					}

				}

			}

		}*/

	/*
	 * for (int i = index; i <= sheet.getLastRowNum(); i++) { XSSFRow row =
	 * sheet.getRow(i); sheet.removeRow(row);
	 * 
	 * }
	 */

	/*
	 * XSSFDrawing xlsx_drawing = plotsheet.createDrawingPatriarch(); 
	 * Define anchor
	 * points in the plotsheet to position the chart 
	 * XSSFClientAnchor anchor =
	 * xlsx_drawing.createAnchor(0, 0, 0, 0, 0, 5, 10, 15); 
	 * Create the chart object
	 * based on the anchor point 
	 * XSSFChart my_line_chart =
	 * xlsx_drawing.createChart(anchor); 
	 * Define legends for the line chart and set
	 * the position of the legend 
	 * XSSFChartLegend legend =
	 * my_line_chart.getOrCreateLegend(); 
	 * legend.setPosition(LegendPosition.BOTTOM);
	 * Create data for the chart 
	 * LineChartData data =
	 * my_line_chart.getChartDataFactory().createLineChartData(); 
	 * Define chart AXIS
	 * ChartAxis bottomAxis =
	 * my_line_chart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
	 * ValueAxis leftAxis =
	 * my_line_chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
	 * leftAxis.setCrosses(AxisCrosses.AUTO_ZERO); 
	 * Define Data sources for the chart
	 * Set the right cell range that contain values for the chart Pass the worksheet
	 * and cell range address as inputs
	 * 
	 * Cell Range Address is defined as First row, last row, first column, last
	 * column
	 * 
	 * ChartDataSource<Number> xs = DataSources.fromNumericCellRange(plotsheet, new
	 * CellRangeAddress(0, 0, 0, 4)); 
	 * ChartDataSource<Number> ys1 =
	 * DataSources.fromNumericCellRange(plotsheet, new CellRangeAddress(1, 1, 0,
	 * 4)); 
	 * ChartDataSource<Number> ys2 =
	 * DataSources.fromNumericCellRange(plotsheet, new CellRangeAddress(2, 2, 0,
	 * 4)); 
	 * ChartDataSource<Number> ys3 =
	 * DataSources.fromNumericCellRange(plotsheet, new CellRangeAddress(3, 3, 0,
	 * 4)); 
	 * Add chart data sources as data to the chart 
	 * data.addSeries(xs, ys1);
	 * data.addSeries(xs, ys2); 
	 * data.addSeries(xs, ys3); 
	 * Plot the chart with the
	 * inputs from data and chart axis 
	 * my_line_chart.plot(data, new ChartAxis[] {
	 * bottomAxis, leftAxis });
	 * 
	 * // 第一次循环取得所有的行的对象 getLastRowNum()是得到最后一行的索引 for (int i = 0; i <=
	 * sheet.getLastRowNum(); i++) { XSSFRow row = sheet.getRow(i); // 获取每一行的对象 for
	 * (int j = 0; j < row.getLastCellNum(); j++) { XSSFCell cell = row.getCell(j);
	 * // 获取每一行的每一列 int type = cell.getCellType(); // 获取每一个单元格对应的类型 switch (type) {
	 * case XSSFCell.CELL_TYPE_BOOLEAN: // 如果是布尔类型 boolean b =
	 * cell.getBooleanCellValue(); System.out.print(b + "    "); break; case
	 * XSSFCell.CELL_TYPE_NUMERIC: // 如果是数字类型 double d = cell.getNumericCellValue();
	 * // 获取值 System.out.print(d + "    "); break; case XSSFCell.CELL_TYPE_STRING:
	 * // 如果是字符串类型的 String s = cell.getStringCellValue(); System.out.print(s +
	 * "   "); case XSSFCell.CELL_TYPE_BLANK: // 如果是空值 System.out.print("      ");
	 * default: break; } } }
	 * 
	 * XSSFSheet newsheet = workbook.createSheet("final version"); // 创建一个工作表
	 * 
	 * // set the style and format
	 * 
	 * XSSFCellStyle style = workbook.createCellStyle(); // 创建单元格风格对象
	 * style.setAlignment(HorizontalAlignment.CENTER); // 设置水平居中
	 * style.setVerticalAlignment(VerticalAlignment.CENTER); // 设置垂直居中
	 * 
	 * XSSFFont font = workbook.createFont(); // 创建字体的对象 font.setFontName("黑体"); //
	 * 设置字体的样式为黑体 font.setFontHeightInPoints((short) 20); // 设置字体的大小
	 * font.setBold(true); // 设置粗体 font.setItalic(true); // 设置倾斜
	 * font.setColor(HSSFColor.RED.index); // 设置字体的颜色
	 * font.setUnderline(FontUnderline.SINGLE); // 设置下划线 font.setStrikeout(false);
	 * // 设置不带下划线
	 * 
	 * style.setFont(font);
	 * 
	 * XSSFRow row1 = newsheet.createRow(0); // 创建第一个行 XSSFCell cell1 =
	 * row1.createCell(0); // 创建第一行的第一列 cell1.setCellStyle(style); //
	 * 将上面定义的风格设置到这个单元格中，这个是必须有的，否则根本不起作用 cell1.setCellValue("demo"); // 设置单元格的内容
	 */

	// workbook.removeSheetAt(0);

	FileOutputStream stream = new FileOutputStream(
			new File("C:\\Users\\ljh\\Desktop\\newtemplate.xlsx"));workbook.write(stream); // 写入文件
	workbook.close(); // 关闭
	stream.close();
	}

	public static void test() throws Exception {
		File file = new File("C:\\Users\\ljh\\Desktop\\template.xlsx");

		FileInputStream fIP = new FileInputStream(file);
		// Get the workbook instance for XLSX file
		XSSFWorkbook workbook = new XSSFWorkbook(fIP);
		if (file.isFile() && file.exists()) {
			System.out.println("openworkbook.xlsx file open successfully.");
		} else {
			System.out.println("Error to open workbook.xlsx file.");
		}

		XSSFSheet sheet = workbook.getSheetAt(0); // 获取第一个工作表的对象

		XSSFRow row = sheet.getRow(0); // 获取每一行的对象

		XSSFCell cell = row.getCell(0); // 获取每一行的每一列
		int type = cell.getCellType(); // 获取每一个单元格对应的类型
		System.out.print(type);

		workbook.close();

	}

}
