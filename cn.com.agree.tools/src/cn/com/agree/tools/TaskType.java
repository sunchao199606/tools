package cn.com.agree.tools;

public enum TaskType {

	BANKPROBLEM("市场问题"), BANKRISK("市场风险"), LEARN("个人成长"), EXAMPLE("案例积累");

	private TaskType(String description) {
		this.taskTypeDescription = description;
	}

	public String getTaskTypeDescription() {
		return taskTypeDescription;
	}

	private String taskTypeDescription;

}
