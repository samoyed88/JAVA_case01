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

class save {
	int groupcount = 0;// 組別數
	int randomnum = 0; // 隨機數字
	ArrayList<String>[] number;// 學號
	ArrayList<String>[] name;// 姓名
	ArrayList<String> records = new ArrayList<String>();// 抽籤記錄

	save() {
	}

	save(int a) {
		this.groupcount = a;
		number = new ArrayList[groupcount];
		name = new ArrayList[groupcount];
		for (int i = 0; i < groupcount; i++) {
			number[i] = new ArrayList<String>();
			name[i] = new ArrayList<String>();
		}
	}

	void randomnumadd() {
		randomnum++;
	}
}

public class Excal_final {

	public static void main(String[] args) {
		Scanner sc = new Scanner(System.in);// 建立Scanner物件
		save original = read();// 原始資料(不修改)
		save leave = original;// 若有請假、缺席可臨時修改
		save Record = new save();// 記錄用
		Record.groupcount = leave.groupcount;// 組別數(透過excel讀取)
		String inputstr;// 使用者輸入數字或n來指定行為
		int input = 0;
		// 從excel讀取組別學號姓名並建立成陣列
		do {
			show();
			inputstr = sc.next();
			try {
				input = Integer.parseInt(inputstr);
			} catch (NumberFormatException e) {
				if (!inputstr.equalsIgnoreCase("n"))
					System.out.println("輸入錯誤請重新輸入");// 輸入n以外皆顯示重新輸入
				continue;
			}
			switch (input) {
			case 1:
				show(leave);
				break;
			case 2:
				case2(leave);
				break;
			case 3:
				case3(leave, Record);
				break;
			case 4:
				case4(leave, Record);
				break;
			case 5:
				reappear(Record);
				break;
			case 6:
				case6(Record);
				break;
			case 7:
				save(Record);
				break;
			default:
				System.out.println("輸入錯誤請重新輸入");
			}
		} while (!inputstr.equalsIgnoreCase("n"));
		// 關閉Scanner
		sc.close();
	}

	// 讀取組別數並且建構original物件
	public static save read() {
		int group = 0;
		try {
			// 使用Apache POI庫中的XSSFWorkbook類別來建立一個Excel工作簿的物件，並從指定的檔案路徑中讀取檔案內容
			XSSFWorkbook xssfWorkbook = new XSSFWorkbook(
					new FileInputStream("C:\\Users\\befor\\eclipse-workspace\\excel\\src\\excel\\分組名單.xlsx"));
			// 使用xssfWorkbook物件的getSheetAt方法來取得第一個工作表(sheet)的物件，並存入sheet變數中
			// XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
			XSSFSheet sheet = xssfWorkbook.getSheet("工作表2");
			// 使用sheet物件的getLastRowNum方法來取得工作表中的最大行數，並存入maxRow變數中
			int maxRow = sheet.getLastRowNum();
			// 使用for迴圈來遍歷工作表中的每一行(row)，從第1行開始，到最大行數結束，每次遞增1
			for (int row = 1; row <= maxRow; row++) {
				// 讀取組別
				XSSFCell groupcell = sheet.getRow(row).getCell(2);
				// group = Integer.parseInt(groupcell.toString());
				group = (int) groupcell.getNumericCellValue();

			}
			save original = new save(group);
			for (int row = 1; row <= maxRow; row++) {
				// 讀取組別
				XSSFCell groupcell = sheet.getRow(row).getCell(2);
				group = (int) groupcell.getNumericCellValue();
				//group = Integer.parseInt(groupcell.toString());
				// 讀取學號
				XSSFCell numcell = sheet.getRow(row).getCell(0);
				// 讀取姓名
				XSSFCell namecell = sheet.getRow(row).getCell(1);
				original.number[group - 1].add(numcell.toString());
				original.name[group - 1].add(namecell.toString());
			}
			return original;
			// 如果發生IOException異常，則捕捉並印出異常的堆疊追蹤
		} catch (IOException e) {
			e.printStackTrace();
			System.exit(0);
		}
		// 建構original物件
		save original = new save(group);
		return original;
	}

	public static void show(save leave) {
		System.out.println("顯示所有人");
		for (int i = 0; i < leave.name.length; i++) {
			System.out.println("第" + (i + 1) + "組：");
			for (int j = 0; j < leave.name[i].size(); j++) {
				System.out.print(leave.name[i].get(j) + " " + leave.number[i].get(j) + " ");
			}
			System.out.println();
		}
		System.out.println();
	}

	public static void search(save leave, String input) {
		int num = -1;
		for (int i = 0; i < leave.number.length; i++) {
			for (int j = 0; j < leave.number[i].size(); j++) {
				num = leave.number[i].indexOf(input);
				if (num != -1) {
					leave.name[i].remove(num);
					leave.number[i].remove(num);
					System.out.print("成功跳過請假者");
					return;// 結束此方法
				}
			}
		}
		System.out.println("查無此人，請重新輸入");
	}

	public static void random(save leave, save Record) {
		for (int i = 0; i < leave.name.length; i++) {
			int a = leave.randomnum % leave.name[i].size();// 隨機數除小組人數的餘數
			System.out.print("本次抽到的為：");
			System.out.println("第" + (i + 1) + "組" + leave.name[i].get(a));
			Record.records.add(leave.number[i].get(a));
			Record.records.add(leave.name[i].get(a));
		}
		leave.randomnumadd();
	}

	public static void random(save leave, save Record, int num) {
		for (int i = 0; i < leave.name.length; i++) {
			int a = num % leave.name[i].size();// 隨機數除小組人數的餘數
			System.out.print("本次抽到的為：");
			System.out.println("第" + (i + 1) + "組" + leave.name[i].get(a));
			Record.records.add(leave.number[i].get(a));
			Record.records.add(leave.name[i].get(a));
		}
		leave.randomnumadd();
	}

	public static void reappear(save Record) {
		int count = 1, group = 1;
		System.out.println("第" + (count++) + "次");
		for (int i = 0; i < Record.records.size(); i += 2) {
			if (group > Record.groupcount) {
				group = 1;
				System.out.println("第" + (count++) + "次");
			}
			System.out.println("第" + (group++) + "組：" + Record.records.get(i) + Record.records.get(i + 1));
		}
		System.out.println();
	}

	public static void reappear(save Record, int count) {
		int group = 1;
		System.out.println("第" + (count) + "次");
		count = (count - 1) * 2 * Record.groupcount;
		for (int i = count; i < count + 2 * Record.groupcount; i += 2) {
			System.out.println("第" + (group++) + "組：" + Record.records.get(i) + Record.records.get(i + 1));
		}
		System.out.println();

		// 若輸入超過次數會造成報錯
	}

	public static void save(save Record) {
		// 保存文件的位置
		String filepath = "C:\\Users\\befor\\eclipse-workspace\\excel\\src\\excel\\save.xlsx";
		try (Workbook workbook = new XSSFWorkbook()) {
			// 創建新的工作表
			Sheet sheet = workbook.createSheet("students");
			int rowcount = 1, count = 1, group = 1;
			// 創建第一行（標題行）
			Row headerRow = sheet.createRow(0);
			// 寫入標題
			Cell headerCell1 = headerRow.createCell(0);
			headerCell1.setCellValue("組別");
			Cell headerCell2 = headerRow.createCell(1);
			headerCell2.setCellValue("學號");
			Cell headerCell3 = headerRow.createCell(2);
			headerCell3.setCellValue("姓名");
			Row serial0 = sheet.createRow(rowcount++);
			Cell serial0Cell1 = serial0.createCell(0);
			serial0Cell1.setCellValue("第" + (count++) + "次");

			for (int i = 0; i < Record.records.size(); i += 2) {
				if (group > Record.groupcount) {
					group = 1;
					Row serial = sheet.createRow(rowcount++);
					Cell serialCell1 = serial.createCell(0);
					serialCell1.setCellValue("第" + (count++) + "次");
				}
				// 創建行
				Row row = sheet.createRow(rowcount++);
				// 創建單元格
				Cell groupCell = row.createCell(0);
				groupCell.setCellValue("第" + (group++) + "組");
				Cell nameCell = row.createCell(1);
				nameCell.setCellValue(Record.records.get(i));
				Cell numberCell = row.createCell(2);
				numberCell.setCellValue(Record.records.get(i + 1));

			}
			// 將工作簿寫入新的 Excel 文件
			try (FileOutputStream fileOut = new FileOutputStream(filepath)) {
				workbook.write(fileOut);
				System.out.println("新的 Excel 文件已保存成功！\r\n");
			}
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public static void show() {
		System.out.println("輸入1可顯示目前名單");
		System.out.println("輸入2可設定請假者");
		System.out.println("輸入3隨機數字抽籤");
		System.out.println("輸入4指定數字抽籤");
		System.out.println("輸入5重現抽籤");
		System.out.println("輸入6可指定第幾次抽籤名單");
		System.out.println("輸入7可將結果存入excel");
		System.out.println("輸入n退出");
	}

	public static void case2(save leave) {
		System.out.println("請輸入請假或缺席者學號，如果輸入完成請輸入N或n");
		Scanner sc = new Scanner(System.in);
		String str = sc.next();
		while (!str.equalsIgnoreCase("n")) {
			search(leave, str);
			str = sc.next();
		}
		// sc.close();
	}

	public static void case3(save leave, save Record) {
		leave.randomnum = (int) (Math.random() * 10);// 隨機一個1~10的數字
		int count = 1;
		String str;
		Scanner sc = new Scanner(System.in);
		do {
			System.out.println("第" + count + "次抽籤名單");
			random(leave, Record);
			System.out.println("輸入任意字可重抽，N退出");
			str = sc.next();
			count++;
		} while (!str.equalsIgnoreCase("n"));
	}

	public static void case4(save leave, save Record) {
		System.out.println("請輸入一指定數字抽籤");
		Scanner sc = new Scanner(System.in);
		random(leave, Record, sc.nextInt());
	}

	public static void case6(save Record) {
		System.out.println("請輸入數字來顯示第n次的抽籤結果");
		Scanner sc = new Scanner(System.in);
		int time = sc.nextInt();
		reappear(Record, time);
	}
}
