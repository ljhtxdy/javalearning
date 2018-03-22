package firstjava;

import java.awt.Font;
import java.awt.RenderingHints;
import java.io.FileOutputStream;

import org.jfree.chart.ChartColor;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartFrame;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.StandardChartTheme;
import org.jfree.chart.axis.CategoryAxis;
import org.jfree.chart.axis.ValueAxis;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.title.TextTitle;
import org.jfree.data.category.CategoryDataset;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.xy.XYSeries;
import org.jfree.data.xy.XYSeriesCollection;

import java.io.*;
import java.util.*;

import javax.swing.JPanel;

public class graph {

	public JFreeChart createLineChart(String code,List<String> data) {		
		// DefaultCategoryDataset dataset = new DefaultCategoryDataset();
		//创建主题样式  
        StandardChartTheme mChartTheme = new StandardChartTheme("CN");  
        //设置标题字体  
        mChartTheme.setExtraLargeFont(new Font("黑体", Font.BOLD, 20));  
        //设置轴向字体  
        mChartTheme.setLargeFont(new Font("宋体", Font.CENTER_BASELINE, 15));  
        //设置图例字体  
        mChartTheme.setRegularFont(new Font("宋体", Font.CENTER_BASELINE, 15));  
        //应用主题样式  
        ChartFactory.setChartTheme(mChartTheme);  

		XYSeries series = new XYSeries("股票代码："+code);
		for (int i = 0; i < data.size(); i++) {
			float result = Float.parseFloat(data.get(data.size()-i-1));
			series.add(i+1, result);
		}

		XYSeriesCollection dataset = new XYSeriesCollection();
		dataset.addSeries(series);
		
		JFreeChart chart = ChartFactory.createXYLineChart("近100个交易日股价", "日期(天)", "股价(元)", dataset, PlotOrientation.VERTICAL, true,
				true, true);

		try {
			ChartUtilities.saveChartAsPNG(new File("E:\\"+code+".png"), chart, 500, 500);
		} catch (IOException e) {
			e.printStackTrace();
		}
		
        return chart;
	}

}
