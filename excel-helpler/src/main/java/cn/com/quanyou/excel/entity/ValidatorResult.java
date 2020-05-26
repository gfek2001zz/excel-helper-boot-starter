package cn.com.quanyou.excel.entity;

public class ValidatorResult {

    private boolean status;
    private String message;

    public boolean getStatus() {
        return status;
    }

    public void setStatus(boolean status) {
        this.status = status;
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }

    public static ValidatorResult ok() {
        ValidatorResult result = new ValidatorResult();
        result.setStatus(true);
        result.setMessage(null);

        return result;
    }

    public static ValidatorResult error(String message) {
        ValidatorResult result = new ValidatorResult();
        result.setStatus(false);
        result.setMessage(message);

        return result;
    }
}
