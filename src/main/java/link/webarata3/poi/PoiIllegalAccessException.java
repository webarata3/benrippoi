package link.webarata3.poi;

public class PoiIllegalAccessException extends RuntimeException {
    public PoiIllegalAccessException() {
    }

    public PoiIllegalAccessException(String message) {
        super(message);
    }

    public PoiIllegalAccessException(String message, Throwable cause) {
        super(message, cause);
    }

    public PoiIllegalAccessException(Throwable cause) {
        super(cause);
    }

    public PoiIllegalAccessException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
