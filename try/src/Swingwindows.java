import java.awt.event.ActionEvent;  

import java.awt.event.ActionListener;  

import java.io.File;  

import java.awt.Dimension;

  

import javax.swing.JButton;  

import javax.swing.JFileChooser;  

import javax.swing.JFrame;  

import javax.swing.JLabel;

import javax.swing.JPanel;

import javax.swing.filechooser.*;

import java.io.FileInputStream;

import java.io.FileOutputStream;

import java.io.IOException;

import java.util.ArrayList;

import org.apache.poi.ss.usermodel.*;

import org.apache.poi.xssf.usermodel.*;

import java.text.*;
import java.util.Date;
import javax.swing.*;

import javax.imageio.ImageIO;
import javax.swing.ImageIcon;
import java.awt.image.BufferedImage;

public class Swingwindows implements ActionListener{  
	    SimpleDateFormat dateFormatter ;
	    JButton Select=null;  

	    JButton Execute=null;

	    JPanel panel;

	    String filename;
	    JLabel label = null;
	    
	    String time;
	    
	    static String IMG_PATH1="C:\\Users\\Meng Xue\\Pictures\\taiyang.jpg";
	    static String IMG_PATH2="C:\\Users\\Meng Xue\\Pictures\\taiyang.jpg";
	    static String IMG_PATH3="C:\\Users\\Meng Xue\\Pictures\\taiyang.jpg";


	    public static void main(String[] args) throws IOException{  

	        Swingwindows GUI=new Swingwindows();  

	        GUI.swingwindows();

	    }  

	    

	    private enum Actions {

	        SELECT,

	        EXECUTE

	      }

	    

	    public void swingwindows(){  


	    	Swingwindows instance = new Swingwindows();

	    	JFrame frame = new JFrame();
	    	panel = new JPanel();
	    	panel.setLayout(null);
	    	final JLabel timeLabel = new JLabel();
	    	final JLabel timeLabel2 = new JLabel();
	    	
	    	final DateFormat timeFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	        ActionListener timerListener = new ActionListener()
	        {
	            public void actionPerformed(ActionEvent e)
	            {
	                Date date = new Date();
	                time = timeFormat.format(date);
	                timeLabel.setText(time);
	            }
	        };
	        Timer timer1 = new Timer(1000, timerListener);
	        // to make sure it doesn't wait one second at the start
	        timer1.setInitialDelay(0);
	        timer1.start();
	        
	       
	        
			BufferedImage img1=null;
			BufferedImage img2=null;
			BufferedImage img3=null;
			
			try {
				img1 = ImageIO.read(new File(IMG_PATH1));
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			
			
			try {
				img2 = ImageIO.read(new File(IMG_PATH2));
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			
			
			try {
				img3 = ImageIO.read(new File(IMG_PATH3));
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			
			ImageIcon icon1 = new ImageIcon(img1);
			ImageIcon icon2 = new ImageIcon(img2);
			ImageIcon icon3 = new ImageIcon(img3);
			

	        
	        
	        ActionListener Listener = new ActionListener()
	        {

	            public void actionPerformed(ActionEvent e)
	            {
				int index = 0;
				char clockarr[] = time.toCharArray();
				String first=Character.toString(clockarr[11]);
				String second=Character.toString(clockarr[12]);
				String strnumber=first+second;
				int number=Integer.parseInt(strnumber); 
				
				if(number<=17 && number>=12)
				{
					index=2;
				}
				else if(number<12) {
					index=1;
				}
				else if(number>17){
					index=3;
					
				}
	
				switch (index) {
				case 1: {
					timeLabel2.setIcon(icon1);
					break;
				}
				case 2: {
					timeLabel2.setIcon(icon2);
					break;
				}
				case 3: {
					timeLabel2.setIcon(icon3);
					break;
				}
				}
			}
	        };
	        
	        Timer timer2= new Timer(1000, Listener);
	        
	        timer2.setInitialDelay(0);
	        timer2.start();
	    
	        
	        
	    
	   
	    	
	        Select=new JButton("Select"); 

	        Execute=new JButton("Execute"); 
	        Select.setBounds(250, 200, 100, 60);
	        Execute.setBounds(250, 300, 100, 60);
	        timeLabel.setBounds(400, 500, 150, 100);
	        timeLabel2.setBounds(0, 500, 150, 100);
	        
	        label=new JLabel();
	        label.setText("Copyright@DSIMS-Jiahui Liu");
	        label.setBounds(0, 0, 200, 30);

	        Select.setActionCommand(Actions.SELECT.name());

	        Select.addActionListener(instance);

	        Execute.setActionCommand(Actions.EXECUTE.name());

	        Execute.addActionListener(instance);

	        panel.add(Select);

	        panel.add(Execute);  
	        
	        panel.add(label,"centre");
	        panel.add(timeLabel);
	        panel.add(timeLabel2);

	        frame.setContentPane(panel);

			frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

			frame.setBounds(0, 0, 250, 250);

			frame.setSize(600, 600);

			frame.setVisible(true);

	    }  


	    public void actionPerformed(ActionEvent e) {  

	    	

	    	if (e.getActionCommand() == Actions.SELECT.name()) {

	    		// TODO Auto-generated method stub  

	            JFileChooser jfc=new JFileChooser();  

	            jfc.setCurrentDirectory(new File("C:\\Users\\Meng Xue\\Desktop\\"));//���õ�ǰĿ¼

	            jfc.setAcceptAllFileFilterUsed(false); //����ѡ�� �����ļ� 

	            

	            

	            FileFilter filter = new FileNameExtensionFilter("Excel","xlsx");

	            jfc.setFileFilter(filter);  



	            jfc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES );  

	            jfc.showDialog(new JLabel(),"ȷ��(�ɱ�ѡ������^-^)");  

	            File file=jfc.getSelectedFile();  

	            /*if(file.isDirectory()){  

	                System.out.println("�ļ���:"+file.getAbsolutePath());  

	            }else if(file.isFile()){  

	                System.out.println("�ļ�:"+file.getAbsolutePath());  

	            } */ 

	            filename=file.getAbsolutePath();   

	    	    } 

	    	else if (e.getActionCommand() == Actions.EXECUTE.name()) {

	    		try {

					RevisedData(filename);

				} catch (Exception e1) {

					// TODO Auto-generated catch block

					e1.printStackTrace();

				}

	    	    }

	    }
    

	public static void RevisedData(String file_name) throws Exception {



		// read the file



		File file = new File(file_name);



		FileInputStream fIP = new FileInputStream(file);

		// Get the workbook instance for XLSX file

		XSSFWorkbook workbook = new XSSFWorkbook(fIP);

		if (file.isFile() && file.exists()) {

			System.out.println("openworkbook.xlsx file open successfully.");

		} else {

			System.out.println("Error to open workbook.xlsx file.");

		}



		XSSFSheet sheet = workbook.getSheetAt(0); // ��ȡ��һ��������Ķ���

		XSSFSheet newsheet123 = workbook.cloneSheet(0, "newsheet123"); // COPY��һ��������Ķ���



		// delete the front part row 1-14

		for (int i = 0; i <= 13; i++) {

			sheet.shiftRows(1, sheet.getLastRowNum(), -1);

		}



		// delete the back part

		int index = 0;



		for (int i = 0; i <= sheet.getLastRowNum(); i++) {

			XSSFRow row = sheet.getRow(i); // ��ȡÿһ�еĶ���



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

			XSSFRow row = sheet.getRow(i); // ��ȡÿһ�еĶ���

			int number = 0;

			ArrayList<Integer> List = new ArrayList<Integer>();

			int decide = 0;



			for (int j = 0; j <= row.getLastCellNum(); j++) {



				XSSFCell cell = row.getCell(j); // ��ȡÿһ�е�ÿһ��

				if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK) {

					number = number + 1;

					List.add(j);

				}

			}



			XSSFRow lastrow = sheet.getRow(i - 1); // ��ȡÿһ�еĶ���

			

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

		

		

		XSSFSheet newsheet = workbook.cloneSheet(0, "new sheet"); // COPY��һ��������Ķ���



		for (int i = part.get(index1-1); i <= newsheet.getLastRowNum(); i++) {

			XSSFRow row = newsheet.getRow(i);

			if (row != null) {

			newsheet.removeRow(row);

			}

		}
		
		

		for (int i = index1-1; i >= 0; i--) {

			XSSFRow startrow = sheet.getRow(1);

			XSSFRow startrow1 = newsheet.getRow(1);

			XSSFRow finalrow = sheet.getRow(1);

			XSSFRow finalrow1 = newsheet.getRow(1);



			startrow = sheet.getRow(part.get(i));

			startrow1 = newsheet.createRow(part.get(i));



			int start = 0;

			int final_ = 0;

			



			start = part.get(i);



			if (i == 0) {

				finalrow = sheet.getRow(sheet.getLastRowNum());

				finalrow1 = newsheet.createRow(sheet.getLastRowNum());



				final_ = sheet.getLastRowNum();



			} else {



				finalrow = sheet.getRow(part.get(i-1) - 1);

				finalrow1 = newsheet.createRow(part.get(i - 1) - 1);



				final_ = part.get(i - 1) - 1;



			}



			int[] List_=new int[sequence[i].length];



			for(int t=0;t<sequence[i].length;t++)

			{

			List_[t] =sequence[i][t];

			}


			Double[] data=new Double[List_.length];
			
			for(int ii=0;ii<List_.length;ii++) {
			
			XSSFCell currentcell = startrow.getCell(List_[ii]);



			data[ii] = currentcell.getNumericCellValue();
			}



			XSSFRow prerow = sheet.getRow(part.get(i) - 1);





			int number123 = 0;



			
			for (int f = 0; f < prerow.getLastCellNum(); f++) {



				XSSFCell precell = prerow.getCell(f);

				// ��ȡÿһ�е�ÿһ��

				if (precell != null && precell.getCellType() != Cell.CELL_TYPE_BLANK) {
					
					number123 = number123 + 1;

				}

			}
			
			int[] prelist = new int[number123];
			
			int number1234=0;
			
			for (int f = 0; f < prerow.getLastCellNum(); f++) {



				XSSFCell precell = prerow.getCell(f);

				// ��ȡÿһ�е�ÿһ��

				if (precell != null && precell.getCellType() != Cell.CELL_TYPE_BLANK) {
					
					prelist[number1234]=f;
					number1234=number1234+1;

				}

			}
			
			
			
			
			Double[] predata=new Double[number123];
		

			int[] preindex = new int[List_.length];
			          
			for (int m = 0; m < number123; m++) {
				
				predata[m] = prerow.getCell(prelist[m]).getNumericCellValue();
			}
			
			

				
			for (int n = 0; n < List_.length; n++) {
				Double[] minimum=new Double[number123/2];
				int groupnumber=0;
				for(int m = 0; m < number123; m++)
				{
					minimum[groupnumber]=data[n]-predata[m];
					m=m+1;
					groupnumber=groupnumber+1;
				}
				
				int index12=0;
				Double minidata=minimum[0];
				
				for(int m=1;m<number123/2;m++) {
					if(minidata>minimum[m])
					{
						minidata=minimum[m];
						index12=m;			
					}
				}
				
				minimum[index12]=0.00;


				preindex[n] = List_[n]-prelist[index12*2];
				preindex[n+1]=List_[n]-prelist[index12*2];

				n=n+1;
			}
			
			for (int b = start; b <= final_; b++) {

				XSSFRow rowb=newsheet.createRow(b);

				for (int c = 0; c < List_.length; c++) {

					if (sheet.getRow(b).getCell(List_[c]) != null

							&& sheet.getRow(b).getCell(List_[c]).getCellType() != Cell.CELL_TYPE_BLANK) {

						Double value = sheet.getRow(b).getCell(List_[c]).getNumericCellValue();

						rowb.createCell(List_[c] -preindex[c]).setCellValue(value);

					}

				}

			}
			
			for (int b = start; b <= final_; b++) {
				XSSFRow rowb=sheet.getRow(b);
				sheet.removeRow(rowb);	
			}
			
			for(int b = start; b <= final_; b++) {
				XSSFRow rowb=sheet.createRow(b);
				XSSFRow newrowb=newsheet.getRow(b);
				for(int c=0;c<newrowb.getLastCellNum();c++) {
					if (newsheet.getRow(b).getCell(c) != null

							&& newsheet.getRow(b).getCell(c).getCellType() != Cell.CELL_TYPE_BLANK) {

						Double value = newsheet.getRow(b).getCell(c).getNumericCellValue();

						rowb.createCell(c).setCellValue(value);
					
				}
				
			}
				
			}
			
			i=0;


		}

		XSSFSheet finalsheet=workbook.createSheet("Revised Data");

		XSSFRow firstrow=newsheet.getRow(0);

		XSSFRow firstrow_=finalsheet.createRow(0);

		int column=firstrow.getLastCellNum();

		int j=0;

		for(int i=0;i<column;i++) {

			if(i%3!=0) {

			String first1= firstrow.getCell(i).getStringCellValue();

			firstrow_.createCell(j).setCellValue(first1);

			j=j+1;

			}			

		}

		

		for(int k=1;k<newsheet.getLastRowNum();k++) {

		finalsheet.createRow(k);}

		

		int mm=0;

		for(int i=0;i<column;i++) {

			if(i%3!=0) {

			for(int k=1;k<newsheet.getLastRowNum();k++) {

				if(newsheet.getRow(k).getCell(i)!= null

						&& newsheet.getRow(k).getCell(i).getCellType() != Cell.CELL_TYPE_BLANK) {

			Double data= newsheet.getRow(k).getCell(i).getNumericCellValue();

			finalsheet.getRow(k).createCell(mm).setCellValue(data);

			}	

			}

			mm=mm+1;

			}

		}

		

		XSSFSheet newsheet1234 = workbook.cloneSheet(1, "Raw Data"); 

		workbook.removeSheetAt(0);

		workbook.removeSheetAt(0);

		workbook.removeSheetAt(0);



	FileOutputStream stream = new FileOutputStream(

			new File(file_name));workbook.write(stream); // д���ļ�

	workbook.close(); // �ر�

	stream.close();

	}
}

	



  




