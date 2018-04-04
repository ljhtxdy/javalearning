package firstjava;

import javax.swing.*;
import java.awt.event.*;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.StandardChartTheme;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.data.xy.XYSeries;
import org.jfree.data.xy.XYSeriesCollection;
import org.jsoup.Jsoup;
import org.jsoup.nodes.*;
import org.jsoup.select.Elements;
import java.io.*;
import java.util.*;
import java.awt.Component;
import java.awt.Font;
import javax.swing.JPanel;

public class Stockanalysis implements ActionListener {
	JButton button;
	JTextField textField;
	JLabel label;
	JLabel label1;
	String code_;
	ChartPanel chart;
	JPanel panel;
	String disindex;

	public static void main1(String[] args) throws IOException {

		Stockanalysis gui = new Stockanalysis();
		gui.window();

	}

	public static Object[] create() throws IOException { // create the stock codes

		List<String> codes = new ArrayList<>();
		List<String> sh = new ArrayList<>();
		List<String> sz = new ArrayList<>();
		sh.add("SS");
		sz.add("SZ");

		try {

			String html = "http://quote.eastmoney.com/stocklist.html";
			Document doc = Jsoup.connect(html).get();

			Elements useful_uls = doc.getElementsByClass("qox");

			List<Element> uls = new ArrayList<>();

			for (Element use : useful_uls) {
				Elements i = use.getElementsByTag("ul");

				for (Element j : i) {
					uls.add(j);
				}

			}

			for (Element ul : uls) {
				Elements lis = ul.getElementsByTag("li");
				for (Element li : lis) {
					String linkText = li.text();
					codes.add(linkText);
				}
			}

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		int dexofsz = 0;

		for (int i = 0; i < codes.size(); i++) {
			String j = codes.get(i);
			if (j.equals("ƽ������(000001)") == true) {
				dexofsz = i;
				break;
			}
		}

		for (int i = 0; i < dexofsz; i++) {
			String j = codes.get(i);
			int index1 = j.indexOf("(");
			int index2 = j.indexOf(")");
			String code = j.substring(index1 + 1, index2);
			sh.add(code);
		}

		for (int k = dexofsz; k < codes.size(); k++) {
			String j = codes.get(k);
			int index1 = j.indexOf("(");
			int index2 = j.indexOf(")");
			String code = j.substring(index1 + 1, index2);
			sz.add(code);
		}

		File file1 = new File("E:\\SH.txt");// Ҫд����ı��ļ�
		if (!file1.exists()) {// ����ļ������ڣ��򴴽����ļ�
			file1.createNewFile();
		}

		File file2 = new File("E:\\SZ.txt");// Ҫд����ı��ļ�
		if (!file2.exists()) {// ����ļ������ڣ��򴴽����ļ�
			file2.createNewFile();
		}

		FileWriter writer1 = new FileWriter(file1);// ��ȡ���ļ��������

		for (int i = 0; i < sz.size(); i++) {
			writer1.write(sz.get(i));// д����
			writer1.write("\r\n");
		}
		writer1.flush();// ��ջ�������������������������д���ļ���
		writer1.close();// �ر��������ʩ����Դ

		FileWriter writer2 = new FileWriter(file2);// ��ȡ���ļ��������

		for (int i = 0; i < sh.size(); i++) {
			writer2.write(sh.get(i));// д����
			writer2.write("\r\n");
		}
		writer2.flush();// ��ջ�������������������������д���ļ���
		writer2.close();// �ر��������ʩ����Դ

		Object[] objArr = new Object[2];
		objArr[0] = sh; // ���ص�һ��ֵ
		objArr[1] = sz;// ���ص�һ��ֵ
		return objArr;

	}

	public static List<String> run(String stock, String where) { // create detailed curve

		List<String> data = new ArrayList<>();

		try {

			String html1 = "https://hk.finance.yahoo.com/quote/";
			String html = html1 + stock + "." + where + "/history?p=" + stock + "." + where;
			Document doc = Jsoup.connect(html).get();
			Elements trs = doc.getElementsByClass("BdT Bdc($c-fuji-grey-c) Ta(end) Fz(s) Whs(nw)");

			List<Element> tr_ = new ArrayList<>();

			for (Element tr : trs) {
				tr_.add(tr);
			}

			for (Element tdss : tr_) {
				int i = 0;
				List<Element> td_ = new ArrayList<>();

				Elements tds = tdss.getElementsByClass("Py(10px) Pstart(10px)");
				for (Element td : tds) {
					i = i + 1;
					td_.add(td);
				}

				if (i == 6) {
					String content = td_.get(4).text();
					data.add(content);
				}
			}

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return data;
	}

	public static String dis(String codeindex) throws IOException {
		List<String> str1 = (List) create()[0]; // warning of type conversion
		int j = 0;
		for (String i : str1) {
			if (i.equals(codeindex)==true) {
				j = 1;
				break;
			}
		}

		List<String> str2 = (List) create()[1]; // warning of type conversion
		for (String i : str2) {
			if (i.equals(codeindex)==true) {
				j = 2;
				break;
			}
		}

		if (j == 2) {
			return "sz";
		} else if (j == 1) {
			return "sh";
		} else {
			return "No";
		}


	}

	public void window() {
		JFrame frame = new JFrame();
		Font font = new Font("Serief", Font.ITALIC + Font.BOLD, 28);// ��������

		panel = new JPanel();

		label = new JLabel();
		label1 = new JLabel();
		button = new JButton("ȷ��");
		textField = new JTextField("", 10);
		textField.addKeyListener(new KeyAdapter() {
			public void keyTyped(KeyEvent e) {
				int keyChar = e.getKeyChar();
				if (keyChar >= KeyEvent.VK_0 && keyChar <= KeyEvent.VK_9) { // ֻ������������

				} else {
					e.consume(); // �ؼ������ε��Ƿ�����
				}
			}
		});

		label.setFont(font);// ���ñ�ǩ����
		label.setText("�ȴ��û�������Ϣ");// ����Ĭ����ʾ����
		label1.setText("�������Ʊ����:(����:000001)");
		
		chart=new ChartPanel(null);

		panel.add(label1);
		panel.add(textField, "centre");
		panel.add(button);
		panel.add(label);

		button.addActionListener(this);

		frame.setContentPane(panel);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.setBounds(100, 100, 250, 250);
		frame.setSize(800, 600);
		frame.setVisible(true);
	}

	public void actionPerformed(ActionEvent event) {
		if (event.getSource() == button) {// �жϴ���Դ�Ƿ�Ϊ��ť
			label.setText(textField.getText());// ���ı��������õ���ǩ
			code_ = textField.getText();
			try {
				disindex = dis(code_);
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			if (disindex.equals("No")==false) {
				panel.remove(chart);
				button.setText("ȷ��");
				List<String> Data = run(code_, disindex);
				chart = new ChartPanel(createLineChart(code_, Data));
				
				panel.add(chart, "centre");
				panel.repaint();
			} else {
				button.setText("��������ȷ�Ĺ�Ʊ����!");
			}

		}
	}
	
	public static JFreeChart createLineChart(String code,List<String> data) {		
		//����������ʽ  
        StandardChartTheme mChartTheme = new StandardChartTheme("CN");  
        //���ñ�������  
        mChartTheme.setExtraLargeFont(new Font("����", Font.BOLD, 20));  
        //������������  
        mChartTheme.setLargeFont(new Font("����", Font.CENTER_BASELINE, 15));  
        //����ͼ������  
        mChartTheme.setRegularFont(new Font("����", Font.CENTER_BASELINE, 15));  
        //Ӧ��������ʽ  
        ChartFactory.setChartTheme(mChartTheme);  

		XYSeries series = new XYSeries("��Ʊ���룺"+code);
		for (int i = 0; i < data.size(); i++) {
			float result = Float.parseFloat(data.get(data.size()-i-1));
			series.add(i+1, result);
		}

		XYSeriesCollection dataset = new XYSeriesCollection();
		dataset.addSeries(series);
		
		JFreeChart chart = ChartFactory.createXYLineChart("��"+data.size()+"�������չɼ�", "����(��)", "�ɼ�(Ԫ)", dataset, PlotOrientation.VERTICAL, true,
				true, true);

		try {
			ChartUtilities.saveChartAsPNG(new File("E:\\"+code+".png"), chart, 500, 400);
		} catch (IOException e) {
			e.printStackTrace();
		}
		
        return chart;
	}

}
