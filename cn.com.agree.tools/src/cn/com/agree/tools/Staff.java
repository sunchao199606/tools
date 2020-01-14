package cn.com.agree.tools;

import java.util.*;

public class Staff {

	private String name;

	private Map<String, Task> taskMap;

	private StaffState staffState;

	public Staff(String name) {
		this.name = name;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public StaffState getStaffState() {
		return staffState;
	}

	public void setStaffState(StaffState staffState) {
		this.staffState = staffState;
	}

	public Map<String, Task> getTaskMap() {
		return taskMap;
	}

	public void setTaskMap(Map<String, Task> taskMap) {
		this.taskMap = taskMap;
	}

}
