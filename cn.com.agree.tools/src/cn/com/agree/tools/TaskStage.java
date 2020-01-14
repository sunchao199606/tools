package cn.com.agree.tools;

public enum TaskStage {

	GETLOGGING("正在取日志"), SOLVING("正在分析"), TESTING("待测试"), SOLVED("已解决");

	private TaskStage(String title) {
		this.title = title;
	}

	public String getTitle() {
		return title;
	}

	private String title;
}
