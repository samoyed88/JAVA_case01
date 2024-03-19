package excel;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.FileOutputStream;
import java.util.*;

public class excelcase01_2 {

	public static void main(String[] args) {
		// 建立Scanner物件
		Scanner sc = new Scanner(System.in);
		// 宣告ArrayList
		ArrayList<String>[] number = new ArrayList[13];
		ArrayList<String>[] name = new ArrayList[13];
		ArrayList<String>[] number2 = new ArrayList[13];
		ArrayList<String>[] name2 = new ArrayList[13];
		for (int i = 0; i < 13; i++) {
			number[i] = new ArrayList<String>();
			number2[i] = new ArrayList<String>();
			name[i] = new ArrayList<String>();
			name2[i] = new ArrayList<String>();
		}
		//從excel讀取組別學號姓名並建立成陣列
		read(name, name2, number, number2);
		// 顯示所有人
		show(name2, number2);
		// 請假者刪除(不會抽到)
		System.out.println("請輸入請假或缺席者學號(可連續輸入)，如果輸入完成請輸入N或n");
		String str = sc.next();
		while (!str.equalsIgnoreCase("n")) {
			search(number2, str, name2);
			str = sc.next();
		}
		// 顯示所有人
		show(name2, number2);
		// 抽籤系統
		int count = 1;//宣告變數count_計算次數用
		int num = (int) (Math.random() * 10);// 隨機一個1~10的數字
		int num2=num;
		do {
			System.out.println("第" + count + "次抽籤名單");
			random(name2, ++num);
			System.out.println("輸入任意字可重抽，N退出");
			str = sc.next();
			count++;
		} while (!str.equalsIgnoreCase("n"));
		// 最原始的名單
		show(name, number);
		//儲存到excel
		input(name2,number2,++num2,count);
		System.out.println("如需要重新查看全部抽到名單請輸入re或輸入數字來取得第n次的名單，輸入其他按鍵則結束");
		String strinput=sc.next();
		if(strinput.equalsIgnoreCase("re"))reappear(num2,count,name2,number2);
		if(Character.isDigit(strinput.charAt(0)))random(name2,num2+(Integer.parseInt(strinput)-1));
		//關閉Scanner
		sc.close();
	}
	
	//讀取excel並導入ArrayList內
	public static void read(ArrayList<String>[] name, ArrayList<String>[] name2, ArrayList<String>[] number,
			ArrayList<String>[] number2) {
		int count = 0;
		try {
			// 使用Apache POI庫中的XSSFWorkbook類別來建立一個Excel工作簿的物件，並從指定的檔案路徑中讀取檔案內容
			XSSFWorkbook xssfWorkbook = new XSSFWorkbook(
					new FileInputStream("C:\\Users\\befor\\OneDrive - Ming Chuan University\\Documents\\student0220.xlsx"));
			// 使用xssfWorkbook物件的getSheetAt方法來取得第一個工作表(sheet)的物件，並存入sheet變數中
			XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
			// 使用sheet物件的getLastRowNum方法來取得工作表中的最大行數，並存入maxRow變數中
			int maxRow = sheet.getLastRowNum();
			// 使用for迴圈來遍歷工作表中的每一行(row)，從第0行開始，到最大行數結束，每次遞增1
			for (int row = 0; row <= maxRow; row++) {
				// 讀取學號
				XSSFCell numcell = sheet.getRow(row).getCell(1);
				number[count / 4].add(numcell.toString());
				number2[count / 4].add(numcell.toString());
				// 讀取姓名
				XSSFCell namecell = sheet.getRow(row).getCell(2);
				name[count / 4].add(namecell.toString());
				name2[count / 4].add(namecell.toString());
				count++;
			}
			// 如果發生IOException異常，則捕捉並印出異常的堆疊追蹤
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	// 隨機抽一組名單
	public static void random(ArrayList<String>[] name, int num) {
		for (int i = 0; i < name.length; i++) {
			int a = num % name[i].size();// 隨機數除小組人數的餘數
			System.out.print("本次抽到的為：");
			System.out.println("第" + (i + 1) + "組" + name[i].get(a));
		}
	}

	// 搜尋某學號並且刪除
	public static void search(ArrayList<String>[] number, String input, ArrayList<String>[] name) {
		int num = -1;
		for (int i = 0; i < number.length; i++) {
			for (int j = 0; j < number[i].size(); j++) {
				num = number[i].indexOf(input);
				if (num != -1) {
					name[i].remove(num);
					number[i].remove(num);
					System.out.print("成功跳過請假者");
					return;//結束此方法
				}
			}
		}System.out.println("查無此人，請重新輸入");
	}

	// 顯示所有名單
	public static void show(ArrayList<String>[] name, ArrayList<String>[] number) {
		System.out.println("顯示所有人");
		for (int i = 0; i < name.length; i++) {
			System.out.println("第" + (i + 1) + "組：");
			for (int j = 0; j < name[i].size(); j++) {
				System.out.print(name[i].get(j) + " " + number[i].get(j) + " ");
			}
			System.out.println();
		}
		System.out.println();
	}
	
	//將抽到的名單輸入excel
	public static void input(ArrayList<String>[] name, ArrayList<String>[] number,int num,int count) {
		// 保存文件的位置
		String filepath="C:\\Users\\befor\\OneDrive - Ming Chuan University\\Documents\\input.xlsx";
		try (Workbook workbook = new XSSFWorkbook()) {
            // 創建新的工作表
            Sheet sheet = workbook.createSheet("students");
            int rowcount=1;
         // 創建第一行（標題行）
            Row headerRow = sheet.createRow(0);

            // 寫入標題
            Cell headerCell1 = headerRow.createCell(0);
            headerCell1.setCellValue("學號");

            Cell headerCell2 = headerRow.createCell(1);
            headerCell2.setCellValue("姓名");
            for(int j=1;j<count;j++) {
            	for (int i = 0; i < name.length; i++) {
            		// 創建行
            		Row row = sheet.createRow(rowcount++);
            		// 寫入姓名和學號信息
            		int a = num % name[i].size();	
                    // 創建單元格
                    Cell nameCell = row.createCell(0);
                    nameCell.setCellValue(name[i].get(a));

                    Cell numberCell = row.createCell(1);
                    numberCell.setCellValue(number[i].get(a));
            	}
            	num++;
            	Row test =sheet.createRow(0);
            	Cell testCell1 = test.createCell(0);
            	testCell1.setCellValue("");
            }

            
            // 將工作簿寫入新的 Excel 文件
            try (FileOutputStream fileOut = new FileOutputStream(filepath)) {
                workbook.write(fileOut);
                System.out.println("新的 Excel 文件已保存成功！");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
		
		
	}
	
	public static void reappear(int num,int count,ArrayList<String>[] name, ArrayList<String>[] number) {
		for(int i=1;i<count;i++){
			random(name,num++);
			System.out.println();
		}
	}
	
	

}