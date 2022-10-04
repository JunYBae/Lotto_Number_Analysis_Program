import java.util.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.NumberFormat;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {	
	public static void clearScreen() {
	    for (int i = 0; i < 80; i++)
	      System.out.println("");
	  }
	
	static Scanner s = new Scanner(System.in);
	static int Cumulative_number[] = new int[46];
	static int Recommended_number[] = new int[7];
	static int analysis_number[] = new int[1000];
	static FileInputStream file;
	static XSSFWorkbook workbook;
	static NumberFormat f;
	static int select_menu = 0;
	static int size = 0;
	static int check_index = 0;
	
	static String path = "C:/Users/bjy54/eclipse-workspace/Lotto_Number_Analysis_Program/";	//���� ���� ��� ����
	static String filename = "excel.xlsx";	//���ϸ� ����
	
	public static void analyze() {
		System.out.println("=================== �ζ� ��ȣ �м� ���α׷� �Դϴ� ===================\n");
		System.out.println("   version 1.1   : ���߹�ȣ ��õ ���α׷�                                 \n");
		System.out.println(" ���� �ּ�(�̸�) : " + filename + "                                  \n");	
		System.out.println("======================================================================\n");
		
		try {		
			System.out.println("                         ��ٷ� �ֽʽÿ�....\n"); // ��� �׷��� ǥ��
			System.out.print("[");
			for(int i = 0; i < 69; i++) {
				Thread.sleep(60);
				System.out.print("=");
			}
			System.out.println("]\n\n");
			Thread.sleep(100);
			clearScreen();
		}
		
		catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		while(true) {
			System.out.println("======================================================================\n");
			System.out.println("                       �޴��� ������ �ֽʽÿ�                           \n");
			System.out.println("----------------------------------------------------------------------\n");
			System.out.println("1. ���� ���� ���� ��ȣ ��õ\n");
			System.out.println("2. ���� �⵵ ���� ��ȣ ��õ [�̱���]\n");
			System.out.println("3. �ý��� ����\n");
			System.out.println("----------------------------------------------------------------------\n");
			System.out.print("���� ��ȣ : ");
			select_menu = s.nextInt();
			System.out.println("\n======================================================================\n");
			
			switch(select_menu) {
			case 1: 
				readExcel(path,filename, 1);
				break;
				
			case 2:
				continue;
				/*System.out.print("���� ȸ�� (���� ȸ��) : ");
				int select_times = s.nextInt();
				readExcel(path, filename, select_times);
				break;*/
				
			case 3:
				System.out.println("\n �ý����� �����մϴ�...\n");
				return;
				
			default:
				System.out.println("            ���� �޴� ��ȣ �Դϴ�. �ٽ� ������ �ֽʽÿ�          \n\n");
				continue;
			}

		}
	}
	
	public static void readExcel(String path, String filename, int select_times) {
		for(int i = 0; i < 46; i++)  // ������ȣ �ʱ�ȭ
			Cumulative_number[i] = 0;		
		for(int i = 1; i < 7; i++)
			Recommended_number[i] = 0;
		for(int i = 0; i < 1000; i++)
			analysis_number[i] = 0;
		size = 0;
		check_index = 0;
		
		try {
			System.out.println("���� �о������ ...\n");
			file = new FileInputStream(filename);
			workbook = new XSSFWorkbook(file);
			f = NumberFormat.getInstance();
			f.setGroupingUsed(false);	//������ �ȳ����� ����
			System.out.println("���� ����");	
			
			try {
				Thread.sleep(1000);
			} catch (InterruptedException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			clearScreen();
			
			int sheetNum = workbook.getNumberOfSheets(); //��Ʈ ����
			
			for(int s = 0; s < sheetNum; s++) {
				
				XSSFSheet sheet = workbook.getSheetAt(s);
			
				int rows = sheet.getPhysicalNumberOfRows(); //�� ����
				
				if(select_times == 1) {  // ��ü �⵵
		
					for(int r = 3 ; r < rows ; r++) {
						XSSFRow row = sheet.getRow(r); // �� ����
					
						int cells = row.getPhysicalNumberOfCells(); // ������ ���� �� ����
						
						for(int c = 13 ; c < cells; c++) {
							XSSFCell cell = row.getCell(c); // �� �о����
							int number = (int)cell.getNumericCellValue();
							Cumulative_number[number]++;	
							size++;
						}										
					}
					
				}	
				
				else if(select_times == 2) {
					System.out.println("\n==================================�̱���=================================\n");
				}	
				
				for(int q = 1; q < 46; q++) {     // ���� Ȯ���� ���Ͽ�, �� Ȯ�� ��ŭ �迭�� ����ִ´�.(�Ҽ��� ����)
					int length = (int)(((double)Cumulative_number[q] / size) * 1000);
					for(int u = 0; u < length; u++) {
						analysis_number[check_index++] = q;
					}
				}
			
				
				
				
				for(int i = 1; i < 7; i++) {
					int max_index = 0;
					int max_number = 0;
					for (int j = 1; j < 46; j++) {
						if(max_number < Cumulative_number[j]) {
							max_index = j;
							max_number = Cumulative_number[j];
						}							
					}
					Cumulative_number[max_index] = 0;
					Recommended_number[i] = max_index;
				}
			}
							
			int minIndex = 0;
			 for(int i = 1; i < 6; i++){
				 int min = 1000;
		            for(int j= i ; j < 7; j++){
		                if(min > Recommended_number[j]) {
		                    minIndex = j;
		                    min = Recommended_number[j];
		                }
		            }		            
		            //������
		            int tmp = Recommended_number[i];
		            Recommended_number[i] = Recommended_number[minIndex];
		            Recommended_number[minIndex] = tmp;
		        }
			 				 
			 for(int i = 0; i < analysis_number.length; i++) {  // ��ȣ Ȯ�� �迭 ���� ����
				 int j = (int)(Math.random()*1000);
				 int temp = analysis_number[i];
				 analysis_number[i] = analysis_number[j];
				 analysis_number[j] = temp;
			 }

			System.out.println("\n==================================��÷��ȣ=================================\n");
			
			System.out.println("---------------------------------------------------------------------------\n");
			System.out.println("                                ���� ���� ��ȣ                             \n");
			System.out.print("\t\t");
			for(int i = 1; i < 7; i++)
				System.out.print("[ " +  Recommended_number[i] + " ]\t");
			System.out.println("\n");
			System.out.println("---------------------------------------------------------------------------\n");
			System.out.println("                                  ��õ ��ȣ                                \n");
			for(int i = 1; i <= 5; i++) {				
				for(int j = 1; j <= 6; j++) {
					 int temp = (int)(Math.random()*1000);
					 boolean check = false;
					 
					 for(int q = 1; q < j; q++) { // �ߺ� �� üũ
							if(Recommended_number[q] == analysis_number[temp])
								check = true;
						 }
					 
					 if (analysis_number[temp] == 0 || check) { // �ߺ��̰ų�, ���� 0�Ͻ� �ٽ� ����
						 j--;
						 continue;						 
					 }
					 						 					 
					 Recommended_number[j] = analysis_number[temp];	 
					 analysis_number[temp] = 0;
				}
				
				int min_Index = 0; 		
				 for(int k = 1; k < 6; k++){		// ���� ����
					 int min = 1000;
			            for(int j= k ; j < 7; j++){
			                if(min > Recommended_number[j]) {
			                    min_Index = j;
			                    min = Recommended_number[j];
			                }
			            }		            
			            //������
			            int tmp = Recommended_number[k];
			            Recommended_number[k] = Recommended_number[min_Index];
			            Recommended_number[min_Index] = tmp;
			        }
				
				System.out.print("\t\t");
				for(int k = 1; k < 7; k++) {
					System.out.print("[ " +  Recommended_number[k] + " ]\t");
				}
				System.out.println("\n");
			}
			
			System.out.println("\n---------------------------------------------------------------------------\n");
			System.out.println("                                               [�޴��� �����Ϸ��� ENTER]");
			String a = s.nextLine();
			String b = s.nextLine();
			System.out.println("===========================================================================\n");
			clearScreen();
			
			file.close();
		}
	
		catch(FileNotFoundException e) {
			System.out.println("������ �������� �ʽ��ϴ�. �ٽ� Ȯ�����ּ���.");
		}
		
		catch(IOException e){
			System.out.println("���� ���� ����");
		}							
	}
	
	public static void main(String[] args) {
		
		analyze();
		
	}
}