package firstjava;


import java.awt.event.ActionEvent;  
import java.awt.event.ActionListener;  
import java.io.File;  
  
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

public class Swingwindows implements ActionListener{  
    JButton Select=null;  
    JButton Execute=null;
    JPanel panel;
    String filename;

    
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
        Select=new JButton("Select"); 
        Execute=new JButton("Execute"); 
        Select.setActionCommand(Actions.SELECT.name());
        Select.addActionListener(instance);
        Execute.setActionCommand(Actions.EXECUTE.name());
        Execute.addActionListener(instance);
        panel.add(Select);
        panel.add(Execute);  
        frame.setContentPane(panel);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.setBounds(100, 100, 250, 250);
		frame.setSize(800, 600);
		frame.setVisible(true);
    }  
    
    public void actionPerformed(ActionEvent e) {  
    	
    	if (e.getActionCommand() == Actions.SELECT.name()) {
    		// TODO Auto-generated method stub  
            JFileChooser jfc=new JFileChooser();  
            jfc.setCurrentDirectory(new File("D:\\"));//设置当前目录
            jfc.setAcceptAllFileFilterUsed(false); //禁用选择 所有文件 
            
            
            FileFilter filter = new FileNameExtensionFilter("Excel","xlsx");
            jfc.setFileFilter(filter);  

            jfc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES );  
            jfc.showDialog(new JLabel(),"确定(可别选错了呦^-^)");  
            File file=jfc.getSelectedFile();  
            /*if(file.isDirectory()){  
                System.out.println("文件夹:"+file.getAbsolutePath());  
            }else if(file.isFile()){  
                System.out.println("文件:"+file.getAbsolutePath());  
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

		XSSFSheet sheet = workbook.getSheetAt(0); // 获取第一个工作表的对象
		XSSFSheet newsheet123 = workbook.cloneSheet(0, "Raw_data_"); // COPY第一个工作表的对象

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
		
		
		XSSFSheet newsheet = workbook.cloneSheet(0, "raw_data"); // COPY第一个工作表的对象

		for (int i = part.get(index1-1); i <= newsheet.getLastRowNum(); i++) {
			XSSFRow row = newsheet.getRow(i);
			if (row != null) {
			newsheet.removeRow(row);
			}
		}
		
		
		
		int extra=0;		
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

			ArrayList<Integer> List_ = new ArrayList<Integer>();

			for(int t=0;t<sequence[i].length;t++)
			{
			List_.add (sequence[i][t]);
			}

			XSSFCell firstcell = startrow.getCell(List_.get(0));

			Double data = firstcell.getNumericCellValue();

			XSSFRow prerow = sheet.getRow(part.get(i) - 1);

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

			int delta = List_.get(0) - prelist.get(preindex)+extra;

			for (int b = start; b <= final_; b++) {
				XSSFRow rowb=newsheet.createRow(b);
				for (int c = 0; c < List_.size(); c++) {

					if (sheet.getRow(b).getCell(List_.get(c)) != null
							&& sheet.getRow(b).getCell(List_.get(c)).getCellType() != Cell.CELL_TYPE_BLANK) {
						Double value = sheet.getRow(b).getCell(List_.get(c)).getNumericCellValue();
						rowb.createCell(List_.get(c) - delta).setCellValue(value);
					}

				}

			}
			
			extra=delta;

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
		
		int m=0;
		for(int i=0;i<column;i++) {
			if(i%3!=0) {
			for(int k=1;k<newsheet.getLastRowNum();k++) {
				if(newsheet.getRow(k).getCell(i)!= null
						&& newsheet.getRow(k).getCell(i).getCellType() != Cell.CELL_TYPE_BLANK) {
			Double data= newsheet.getRow(k).getCell(i).getNumericCellValue();
			finalsheet.getRow(k).createCell(m).setCellValue(data);
			}	
			}
			m=m+1;
			}
		}
		
		XSSFSheet newsheet1234 = workbook.cloneSheet(1, "Raw Data"); 
		workbook.removeSheetAt(0);
		workbook.removeSheetAt(0);
		workbook.removeSheetAt(0);


	FileOutputStream stream = new FileOutputStream(
			new File(file_name));workbook.write(stream); // 写入文件
	workbook.close(); // 关闭
	stream.close();
	}
	

  
} 