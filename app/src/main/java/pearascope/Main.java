package pearascope;

import java.time.Year;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) {
        String year = String.valueOf(Year.now().getValue()); // default to current year
        boolean enableMonitoring = false;
        boolean generateRawDump = false;

        List<String> filePathList = new ArrayList<>();
        for (String arg : args) {
            if (arg.toLowerCase().startsWith("-year:")) {
                year = arg.substring(6);
            } else if (arg.toLowerCase().startsWith("-monitor")) {
                enableMonitoring = true;
            } else if (arg.toLowerCase().startsWith("-raw")) {
                generateRawDump = true;
            } else if (arg.toLowerCase().startsWith("-help") || arg.toLowerCase().startsWith("-h")) {
                System.out.println("Usage: java -jar Pearascope.jar [-year:<year>] [-monitor] [-raw] [<file1> <file2> ...]");
                System.out.println("  -year:<year>   Specify the year for the processor (default is current year).");
                System.out.println("  -monitor       Monitor the input folder for new logs.");
                System.out.println("  -raw           Generate raw dump of logs (cannot be used with -monitor).");
                System.out.println("  <file1>...     One or more paths to log files to process (will be ignored if using -monitor).");
                return;
            } else {
                filePathList.add(arg);
            }
        }

        String[] filePaths = filePathList.toArray(new String[0]);

        String className = "pearascope.Pearascope" + year;
        try {
            Class<?> processorClass = Class.forName(className);
            Object instance = processorClass.getDeclaredConstructor().newInstance();

            if (instance instanceof PearascopeProcessor) {
                System.out.println("Using processor for year " + year);
                PearascopeProcessor processor = (PearascopeProcessor) instance;
                processor.run(enableMonitoring, generateRawDump, filePaths);
            } else {
                System.err.println("Class " + className + " does not implement PearascopeProcessor.");
            }
        } catch (ClassNotFoundException e) {
            System.err.println("No processor found for year " + year + ". Class not found: " + className);
        } catch (Exception e) {
            System.err.println("Error instantiating processor for year " + year);
            e.printStackTrace();
        }
    }
}
