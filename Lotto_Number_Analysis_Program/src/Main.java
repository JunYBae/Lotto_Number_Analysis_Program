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
	static FileInputStream file;
	static XSSFWorkbook workbook;
	static NumberFormat f;
	static int select_menu = 0;
	
	static String path = "C:/Users/bjy54/eclipse-workspace/Lotto_Number_Analysis_Program/";	//파일 경로 설정
	static String filename = "excel.xlsx";	//파일명 설정
	
	
	public static void analyze() {
		System.out.println("=================== 로또 번호 분석 프로그램 입니다 ===================\n");
		System.out.println(" version 1.1 : 다중번호 추천 프로그램                                 \n");
		System.out.println("  파일 주소  : " + path+filename + "                                  \n");	
		System.out.println("======================================================================\n");
		
		try {		
			System.out.println("                         기다려 주십시오....\n"); // 대기 그래픽 표현
			System.out.print("[");
			for(int i = 0; i < 69; i++) {
				Thread.sleep(80);
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
			System.out.println("                       메뉴를 선택해 주십시오                           \n");
			System.out.println("----------------------------------------------------------------------\n");
			System.out.println("1. 역대 다중 출현 번호 추천\n");
			System.out.println("2. 지정 년도 다중 번호 추천 [미구현]\n");
			System.out.println("3. 시스템 종료\n");
			System.out.println("----------------------------------------------------------------------\n");
			System.out.print("선택 번호 : ");
			select_menu = s.nextInt();
			System.out.println("\n======================================================================\n");
			
			switch(select_menu) {
			case 1: 
				readExcel(path,filename, 1);
				break;
				
			case 2:
				System.out.print("지정 회차 (누적 회차) : ");
				int select_times = s.nextInt();
				readExcel(path, filename, select_times);
				break;
				
			case 3:
				System.out.println("\n 시스템을 종료합니다...\n");
				return;
				
			default:
				System.out.println("            없는 메뉴 번호 입니다. 다시 선택해 주십시오          \n\n");
				continue;
			}

		}
	}
	
	public static void readExcel(String path, String filename, int select_times) {
		for(int i = 0; i < 46; i++)  // 누적번호 초기화
			Cumulative_number[i] = 0;		
		for(int i = 1; i < 7; i++)
			Recommended_number[i] = 0;
		
		try {
			System.out.println("파일 읽어오는중 ...\n");
			file = new FileInputStream(path + filename);
			workbook = new XSSFWorkbook(file);
			f = NumberFormat.getInstance();
			f.setGroupingUsed(false);	//지수로 안나오게 설정
			System.out.println("파일 열림");	
			
			try {
				Thread.sleep(1000);
			} catch (InterruptedException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			clearScreen();
			
			int sheetNum = workbook.getNumberOfSheets(); //시트 갯수
			
			for(int s = 0; s < sheetNum; s++) {
				
				XSSFSheet sheet = workbook.getSheetAt(s);
			
				int rows = sheet.getPhysicalNumberOfRows(); //행 갯수
				
				if(select_times == 1) {  // 전체 년도
		
					for(int r = 3 ; r < rows ; r++) {
						XSSFRow row = sheet.getRow(r); // 행 지정
					
						int cells = row.getPhysicalNumberOfCells(); // 지정된 행의 열 갯수
						
						for(int c = 13 ; c < cells; c++) {
							XSSFCell cell = row.getCell(c); // 셀 읽어오기
							int number = (int)cell.getNumericCellValue();
							Cumulative_number[number]++;							
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
				
				else if(select_times == 2) {
					System.out.println("\n==================================미구현=================================\n");
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
		            //스와프
		            int tmp = Recommended_number[i];
		            Recommended_number[i] = Recommended_number[minIndex];
		            Recommended_number[minIndex] = tmp;
		        }
			
			System.out.println("\n==================================추첨번호=================================\n");
			
			System.out.println("---------------------------------------------------------------------------\n");
			System.out.print("\t\t");
			for(int i = 1; i < 7; i++)
				System.out.print("[ " +  Recommended_number[i] + " ] ");
			System.out.println();
			System.out.println("\n---------------------------------------------------------------------------\n");
			System.out.println("                                               [메뉴로 복귀하려면 ENTER]");
			String a = s.nextLine();
			String b = s.nextLine();
			System.out.println("===========================================================================\n");
			clearScreen();
			
			file.close();
		}
	
		catch(FileNotFoundException e) {
			System.out.println("파일이 존재하지 않습니다. 다시 확인해주세요.");
		}
		
		catch(IOException e){
			System.out.println("파일 열기 오류");
		}							
	}
	
	public static void main(String[] args) {
		
		analyze();
		
	}
}