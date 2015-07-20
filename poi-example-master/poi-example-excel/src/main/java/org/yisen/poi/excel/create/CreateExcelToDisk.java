package org.yisen.poi.excel.create;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class CreateExcelToDisk {

	/**
	 * 手工构建一个简单格式的Excel
	 */
	private static List<Student> getStudent() throws Exception {
		List<Student> list = new ArrayList<Student>();
		SimpleDateFormat df = new SimpleDateFormat("yyyy-mm-dd");

		Student user1 = new Student(1, "张三", 16, df.parse("1997-03-12"));
		Student user2 = new Student(2, "李四", 17, df.parse("1996-08-12"));
		Student user3 = new Student(3, "王五", 26, df.parse("1985-11-12"));
		Student user4 = new Student(4, "李四", 15, df.parse("1896-08-15"));
		Student user5 = new Student(5, "李5", 16, df.parse("1946-08-22"));
		Student user6 = new Student(6, "李6", 17, df.parse("1966-05-14"));

		list.add(user1);
		list.add(user2);
		list.add(user3);
		list.add(user4);
		list.add(user5);
		list.add(user6);

		return list;
	}

	public static void createSimpleExcelToDisk() throws Exception {
		// 第一步，创建一个webbook，对应一个Excel文件
		HSSFWorkbook wb = new HSSFWorkbook();
		// 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
		HSSFSheet sheet = wb.createSheet("学生表一");
		// 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制short
		HSSFRow row = sheet.createRow((int) 0);
		// 第四步，创建单元格，并设置值表头 设置表头居中
		HSSFCellStyle style = wb.createCellStyle();
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 创建一个居中格式

		HSSFCell cell = row.createCell(0);
		cell.setCellValue("学号");
		cell.setCellStyle(style);
		cell = row.createCell(1);
		cell.setCellValue("姓名");
		cell.setCellStyle(style);
		cell = row.createCell(2);
		cell.setCellValue("年龄");
		cell.setCellStyle(style);
		cell = row.createCell(3);
		cell.setCellValue("生日");
		cell.setCellStyle(style);

		// 第五步，写入实体数据 实际应用中这些数据从数据库得到，
		List<Student> list = CreateExcelToDisk.getStudent();
		for (int i = 0; i < list.size(); i++) {
			row = sheet.createRow((int) i + 1);
			Student stu = (Student) list.get(i);
			// 第四步，创建单元格，并设置值
			row.createCell(0).setCellValue((double) stu.getId());
			row.createCell(1).setCellValue(stu.getName());
			row.createCell(2).setCellValue((double) stu.getAge());
			cell = row.createCell(3);
			cell.setCellValue(new SimpleDateFormat("yyyy-mm-dd").format(stu
					.getBirth()));
		}

		FileOutputStream fout = new FileOutputStream("./students.xls");
		wb.write(fout);
		wb.close();
		fout.close();

	}

	public static void createMergedRegionExcelToDisk() throws Exception {
		// 第一步，创建一个webbook，对应一个Excel文件
		HSSFWorkbook wb = new HSSFWorkbook();
		// 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
		HSSFSheet sheet = wb.createSheet("学生表一");
		// 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制short
		HSSFRow row = sheet.createRow((int) 0);
		// 第四步，创建单元格，并设置值表头 设置表头居中
		HSSFCellStyle style = wb.createCellStyle();
		// 创建一个居中格式
		style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// 垂直
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);// 水平

		HSSFCell cell = row.createCell(0);
		cell.setCellValue("学号");
		cell.setCellStyle(style);
		cell = row.createCell(1);
		cell.setCellValue("姓名");
		cell.setCellStyle(style);
		cell = row.createCell(2);
		cell.setCellValue("年龄");
		cell.setCellStyle(style);
		cell = row.createCell(3);
		cell.setCellValue("生日");
		cell.setCellStyle(style);

		// 第五步，写入实体数据 实际应用中这些数据从数据库得到，
		List<Student> list = CreateExcelToDisk.getStudent();
		for (int i = 0; i < list.size(); i++) {
			row = sheet.createRow(i + 1);
			Student stu = (Student) list.get(i);
			// 第四步，创建单元格，并设置值
			cell = row.createCell(0);
			cell.setCellValue((double) stu.getId());
			cell.setCellStyle(style);

			cell = row.createCell(1);
			cell.setCellValue(stu.getName());
			cell.setCellStyle(style);

			cell = row.createCell(2);
			cell.setCellValue((double) stu.getAge());
			cell.setCellStyle(style);

			cell = row.createCell(3);
			cell.setCellValue(new SimpleDateFormat("yyyy-mm-dd").format(stu
					.getBirth()));
			cell.setCellStyle(style);

		}

		// 指定合并区域
		// sheet.addMergedRegion(new Region(1,(short)1,1,(short)2));
		sheet.addMergedRegion(new CellRangeAddress(1, 3, 1, 1));

		System.out.println(sheet.getColumnStyle(1));

		FileOutputStream fout = new FileOutputStream(
				"./students-merged-region.xls");
		wb.write(fout);
		wb.close();
		fout.close();

	}

}
