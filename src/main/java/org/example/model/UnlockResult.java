package org.example.model;
import java.io.File;
public class UnlockResult {
    private boolean success;
    private String message;
    private File generatedFile;
    private String exceptionMessage;

    public UnlockResult(boolean success, String message, File generatedFile, String exceptionMessage) {
        this.success = success;
        this.message = message;
        this.generatedFile = generatedFile;
        this.exceptionMessage = exceptionMessage;
    }

    public boolean isSuccess() {
        return success;
    }

    public String getMessage() {
        return message;
    }

    public File getGeneratedFile() {
        return generatedFile;
    }

    public String getExceptionMessage() {
        return exceptionMessage;
    }
}
