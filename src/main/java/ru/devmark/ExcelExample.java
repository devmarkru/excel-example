package ru.devmark;

import ru.devmark.reader.ExcelReader;
import ru.devmark.writer.ExcelWriter;

import java.io.IOException;

public class ExcelExample {

    public static void main(String[] args) throws IOException {
        if (args.length < 2) {
            System.out.println("Please, specify action (read, write) and Excel file location");
            return;
        }
        var action = args[0].toLowerCase();
        var filename = args[1];
        switch (action) {
            case "read":
                var reader = new ExcelReader();
                reader.read(filename);
                break;
            case "write":
                var writer = new ExcelWriter();
                writer.write(filename);
                break;
            default:
                throw new RuntimeException("Unknown action: " + action);
        }
    }
}