package cn.com.agree.tools;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.xmind.core.Core;
import org.xmind.core.CoreException;
import org.xmind.core.IDeserializer;
import org.xmind.core.ISerializer;
import org.xmind.core.ISheet;
import org.xmind.core.ITopic;
import org.xmind.core.IWorkbook;
import org.xmind.core.internal.dom.TopicImpl;

/**
 * excel里面的周报转化为xmind格式
 * 
 * @author sunchao
 *
 */
public class WeeklyTaskUtil {

	private static final String EXCEL_PATH = "E:\\89754\\weixin\\WeChat Files\\s_cgogogo\\FileStorage\\File\\2020-01\\周报_AB4.0市场支持与研发周报20200106-20200112.xlsx";
	private static final String OUTPUT_DIR = "E:\\";

	/**
	 * 周报excel转xmind
	 */
	@Test
	public void excelToXmind() {
		// 1 加载excel数据
		Workbook excelWorkbook = getWorkbook(EXCEL_PATH);

		Map<String, Staff> staffMap = parseDataFromExcelWorkbook(excelWorkbook);
		// 2 反序列化模板文件
		IDeserializer deserializer = Core.getWorkbookBuilder().newDeserializer();
		try {
			String templateXmindPath = WeeklyTaskUtil.class.getResource("./resources/template.xmind").getPath();
			InputStream fileInputStream = new FileInputStream(templateXmindPath);
			deserializer.setInputStream(fileInputStream);
			deserializer.deserialize(null);
		} catch (IllegalStateException | IOException | CoreException e1) {
			e1.printStackTrace();
		}
		// 3 对模板文件重新进行编辑
		IWorkbook xmindWorkbook = deserializer.getWorkbook();
		putDataInXmindWorkbook(xmindWorkbook, staffMap);

		// 4 导出目标文件
		String sheetName = excelWorkbook.getSheetAt(0).getSheetName();
		String date = sheetName.substring(0, sheetName.length() - 2);
		String fileName = "AB4.0市场支持与研发周报" + date;
		xmindWorkbook.getPrimarySheet().getRootTopic().setTitleText(fileName);
		ISerializer serializer = Core.getWorkbookBuilder().newSerializer();
		try {
			String outputPath = OUTPUT_DIR + "\\" + fileName + ".xmind";
			OutputStream outputStream = new FileOutputStream(outputPath);
			serializer.setOutputStream(outputStream);
			serializer.setWorkbook(xmindWorkbook);
			serializer.serialize(null);
		} catch (IllegalStateException | IOException | CoreException e) {
			e.printStackTrace();
		}
	}

	/**
	 * 
	 * 合并周报
	 */
	@Test()
	public void mergeExcel() {
		// 搜集excel数据
		String excelPath = "C:\\Users\\89754\\Desktop\\个人周报";
		File dir = new File(excelPath);
		Map<String, List<Task>> taskMap = new HashMap<>();
		List<Workbook> workbookList = new ArrayList<>();
		for (File file : dir.listFiles()) {
			Workbook wb = null;
			try (InputStream is = new FileInputStream(file.getPath());) {
				wb = WorkbookFactory.create(is);
			} catch (Exception e) {
				e.printStackTrace();
			} finally {
				if (wb != null) {
					try {
						wb.close();
					} catch (IOException e) {
						e.printStackTrace();
					}
				}
			}
			workbookList.add(wb);
		}
		// 合并数据
		List<Map<String, Staff>> staffMapList = parseDataFromExcelWorkbooks(workbookList);
		staffMapList.forEach(map -> {
			map.entrySet().forEach(entry -> {
				String key = entry.getKey().trim();
				Staff staff = entry.getValue();
				Collection<Task> tasks = staff.getTaskMap().values();
				if (taskMap.containsKey(key)) {
					taskMap.get(key).addAll(tasks);
				} else {
					List<Task> taskList = new ArrayList<>(tasks);
					taskMap.put(key, taskList);
				}
			});
		});

		// 导出到excel
		exportDataToExcel(taskMap);

	}

	/**
	 * 数据导出到excel中
	 * 
	 * @param taskMap
	 */
	private void exportDataToExcel(Map<String, List<Task>> taskMap) {

		Workbook workbook = new XSSFWorkbook();
		// 生成一个表格
		Sheet sheet = workbook.createSheet("2020上半年汇总");
		// 设置表格默认列宽度为15个字节
		sheet.setDefaultColumnWidth((short) 15);
		// 生成表头样式
		CellStyle style = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setBold(true);
		style.setFont(font);
		// 创建表头
		Row row = sheet.createRow(0);
		String[] titles = new String[] { "姓名", "TAPD需求ID", "问题类型", "市场版本", "需求标题" };
		Map<String, Integer> titleOrder = new HashMap<>();
		for (int i = 0; i < 5; i++) {
			Cell cell = row.createCell(i);
			cell.setCellStyle(style);
			String title = titles[i];
			cell.setCellValue(title);
			titleOrder.put(title, i);
		}
		// 写入正文
		AtomicInteger index = new AtomicInteger(1);
		AtomicInteger startRow = new AtomicInteger(1);
		AtomicInteger endRow = new AtomicInteger();
		taskMap.entrySet().forEach(entry -> {
			String name = entry.getKey();
			// System.out.println(name);
			List<Task> tasks = entry.getValue();
			startRow.set(endRow.get() + 1);
			// System.out.println(startRow);
			endRow.set(startRow.get() + tasks.size() - 1);
			// System.out.println(endRow);
			for (Task task : tasks) {
				Row r = sheet.createRow(index.get());
				String id = task.getId() == null ? "" : task.getId();
				if (id.startsWith("unknow")) {
					id = "";
				}
				String type = task.getTaskType() == null ? "" : task.getTaskType().getTaskTypeDescription();
				String market = task.getMarket() == null ? "" : task.getMarket();
				String description = task.getTaskDescription() == null ? "" : task.getTaskDescription();
				Cell nameCell = r.createCell(0);
				nameCell.setCellValue(name);
				Cell idCell = r.createCell(1);
				idCell.setCellValue(id);
				Cell typeCell = r.createCell(2);
				typeCell.setCellValue(type);
				Cell marketCell = r.createCell(3);
				marketCell.setCellValue(market);
				Cell descriptionCell = r.createCell(4);
				descriptionCell.setCellValue(description);
				index.incrementAndGet();
			}
			// System.out.println(endRow);
			// 合并单元格
			if (tasks.size() > 1) {
				CellRangeAddress region = new CellRangeAddress(startRow.get(), endRow.get(), 0, 0);
				sheet.addMergedRegion(region);
			}
		});

		try (FileOutputStream outputStream = new FileOutputStream("D:\\周报汇总.xlsx");) {
			workbook.write(outputStream);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				workbook.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	private static Map<String, Staff> parseDataFromExcelWorkbook(Workbook workbook) {
		// System.out.println("*************************");
		Sheet sheet = workbook.getSheetAt(0);
		// System.out.println(sheet.getSheetName());
		Map<String, Staff> staffMap = new HashMap<String, Staff>();
		int seed = 0;
		AtomicReference<String> cacheName = new AtomicReference<String>("");
		// String cacheName = "";
		for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
			Row row = sheet.getRow(rowIndex);
			if (row.getCell(0) == null) {
				System.out.println(sheet.getSheetName() + "第" + rowIndex++ + "行数据为空");
				continue;
			}
			if (row == null || (row.getCell(0) != null && row.getCell(0).getStringCellValue().equals("姓名"))) {
				System.out.println("过滤掉表头");
				continue;
			}
			// 获取name并缓存
			String name = row.getCell(0).getStringCellValue().split("-")[0];
			if (name != null && !name.isEmpty()) {
				cacheName.set(name);
			}
			staffMap.computeIfAbsent(cacheName.get(), k -> {
				System.out.println("解析" + cacheName.get() + "任务数据..");
				Staff staff = new Staff(k);
				staff.setTaskMap(new HashMap<String, Task>());
				return staff;
			});
			// 获取Id
			Cell c = row.getCell(1);
			String taskId = "";
			if (c != null) {
				try {
					taskId = c.getStringCellValue().trim().replaceAll("[^0-9]", "");
				} catch (IllegalStateException e) {
					System.out.println(cacheName.get() + "任务" + c.getNumericCellValue() + "任务ID为数值类型");
					taskId = String.valueOf((int) c.getNumericCellValue());
				}
			}
			Map<String, Task> taskMap = staffMap.get(cacheName.get()).getTaskMap();
			if (taskId.isEmpty()) {
				System.out.println(cacheName.get() + "任务ID为空");
				// 看下任务标题是否为空
				AtomicReference<String> taskTitle = new AtomicReference<String>("");
				if (row.getCell(4) != null) {
					taskTitle.set(row.getCell(4).getStringCellValue());
				}
				if (!taskTitle.get().isEmpty()) {
					taskId = "unknowID" + (++seed);
					taskMap.computeIfAbsent(taskId, k -> {
						Task task = new Task(k);
						task.setTaskDescription(taskTitle.get());
						return task;
					});
				} else {
					System.out.println("第" + (rowIndex + 1) + "行为无效数据行,过滤掉");
					continue;
				}
			} else {
				taskMap.computeIfAbsent(taskId, k -> new Task(k));
			}

			for (int cellIndex = 2; cellIndex <= row.getLastCellNum(); cellIndex++) {
				Cell cell = row.getCell(cellIndex);
				if (cell == null) {
					System.out.println("第" + (rowIndex + 1) + "行,第" + (cellIndex + 1) + "列数据为空");
					continue;
				}
				switch (cellIndex) {
				// 问题类型
				case 2:
					String taskType = cell.getStringCellValue().trim();
					if (taskType != null) {
						TaskType t = null;
						if ("市场问题".equals(taskType)) {
							t = TaskType.BANKPROBLEM;
						} else if ("市场风险".equals(taskType)) {
							t = TaskType.BANKRISK;
						} else if ("个人成长".equals(taskType)) {
							t = TaskType.LEARN;
						} else if ("案例积累".equals(taskType)) {
							t = TaskType.EXAMPLE;
						} else {
							System.out.println("无法识别的任务类型:" + taskType);
						}
						staffMap.get(cacheName.get()).getTaskMap().get(taskId).setTaskType(t);
					}
					break;
				// 市场版本
				case 3:
					String bankMarket = cell.getStringCellValue().trim();
					if (bankMarket != null) {
						staffMap.get(cacheName.get()).getTaskMap().get(taskId).setMarket(bankMarket);
					}
					break;
				// 需求标题
				case 4:
					String taskTitle = cell.getStringCellValue();
					if (taskTitle != null) {
						staffMap.get(cacheName.get()).getTaskMap().get(taskId).setTaskDescription(taskTitle);
					}
					break;
				// 进度
				case 5:
					String taskStage = cell.getStringCellValue().trim();
					if (taskStage != null) {
						TaskStage t = null;
						if ("正在拿日志".equals(taskStage)) {
							t = TaskStage.GETLOGGING;
						} else if ("正在分析".equals(taskStage)) {
							t = TaskStage.SOLVING;
						} else if ("待测试".equals(taskStage)) {
							t = TaskStage.TESTING;
						} else if ("已解决".equals(taskStage)) {
							t = TaskStage.SOLVED;
						} else {
							System.out.println("无法识别的任务进度:" + taskStage);
						}
						staffMap.get(cacheName.get()).getTaskMap().get(taskId).setTaskStage(t);
					}
					break;
				// 个人/市场状态
				case 6:
					String state = cell.getStringCellValue().trim();
					if (state != null && !state.isEmpty()) {
						StaffState s = null;
						if ("问题不多，比较轻松".equals(state)) {
							s = StaffState.EASY;
						} else if ("问题一般多，在把控之内".equals(state)) {
							s = StaffState.NORMAL;
						} else if ("问题较多，超出把控".equals(state)) {
							s = StaffState.BUSY;
						} else if ("必须要多人解决".equals(state)) {
							s = StaffState.OVERLOAD;
						} else {
							System.out.println("无法识别的个人状态:" + s);
						}
						staffMap.get(cacheName.get()).setStaffState(s);
					}
					break;
				default:
					break;
				}
			}
		}
		return staffMap;
	}

	private static List<Map<String, Staff>> parseDataFromExcelWorkbooks(List<Workbook> workbookList) {
		List<Map<String, Staff>> staffMapList = new ArrayList<>();
		for (Workbook workbook : workbookList) {
			staffMapList.add(parseDataFromExcelWorkbook(workbook));
		}
		return staffMapList;
	}

	private static void putDataInXmindWorkbook(IWorkbook xmindWorkbook, Map<String, Staff> staffMap) {
		ISheet sheet = xmindWorkbook.getPrimarySheet();
		ITopic rootTopic = sheet.getRootTopic();
		// 将所有关联节点删除
		rootTopic.getChildren(TopicImpl.ATTACHED).forEach(t -> {
			rootTopic.remove(t);
		});
		staffMap.values().stream().forEach(staff -> {
			ITopic staffTopic = xmindWorkbook.createTopic();
			staffTopic.setTitleText(staff.getName());
			if (staff.getStaffState() == StaffState.EASY) {
				staffTopic.addMarker("smiley-laugh");
			} else if (staff.getStaffState() == StaffState.NORMAL) {
				staffTopic.addMarker("smiley-embarrass");
			} else if (staff.getStaffState() == StaffState.BUSY) {
				staffTopic.addMarker("smiley-surprise");
			} else if (staff.getStaffState() == StaffState.OVERLOAD) {
				staffTopic.addMarker("smiley-cry");
			} else {
				staffTopic.addMarker("smiley-embarrass");
			}
			for (Task task : staff.getTaskMap().values()) {
				ITopic marketTopic = null;
				String market = task.getMarket() == null ? "" : task.getMarket();
				if (staffTopic.hasChildren(TopicImpl.ATTACHED)) {
					for (ITopic mt : staffTopic.getAllChildren()) {
						if (market.equals(mt.getTitleText())) {
							marketTopic = mt;
						}
					}
				}
				if (marketTopic == null) {
					marketTopic = xmindWorkbook.createTopic();
					staffTopic.add(marketTopic);
				}
				marketTopic.setTitleText(task.getMarket());
				ITopic taskTopic = xmindWorkbook.createTopic();
				taskTopic.setTitleWidth(500);
				String taskId = task.getId();
				if (taskId.startsWith("unknow")) {
					taskTopic.setTitleText(task.getTaskDescription());
				} else {
					taskTopic.setTitleText("[" + task.getId() + "]" + task.getTaskDescription());
				}
				if (task.getTaskType() == TaskType.BANKPROBLEM) {
					taskTopic.addMarker("symbol-question");
				} else if (task.getTaskType() == TaskType.BANKRISK) {
					taskTopic.addMarker("symbol-attention");
				} else if (task.getTaskType() == TaskType.EXAMPLE) {
					taskTopic.addMarker("c_symbol_pen");
				} else if (task.getTaskType() == TaskType.LEARN) {
					taskTopic.addMarker("c_symbol_exercise");
				} else {
					taskTopic.addMarker("c_symbol_exercise");
				}
				marketTopic.add(taskTopic);
			}
			rootTopic.add(staffTopic);
		});
	}

	private static Workbook getWorkbook(String filepath) {
		Workbook wb = null;
		try (InputStream is = new FileInputStream(filepath);) {
			wb = WorkbookFactory.create(is);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (wb != null) {
				try {
					wb.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return wb;
	}

}
