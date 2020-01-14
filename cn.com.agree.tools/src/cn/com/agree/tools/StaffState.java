package cn.com.agree.tools;

public enum StaffState {

    EASY("问题不多，比较轻松"), NORMAL("问题一般多，在把控之内"), BUSY("问题较多，超出把控"), OVERLOAD(
            "必须要多人解决");

    private String description;

    private StaffState(String description) {
        this.description = description;
    }

    public String getDescription() {
        return description;
    }

}
