package pearascope;

import java.io.File;
import java.io.IOException;
import java.nio.file.FileSystems;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardWatchEventKinds;
import java.nio.file.WatchEvent;
import java.nio.file.WatchKey;
import java.nio.file.WatchService;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import edu.wpi.first.util.datalog.DataLogReader;
import edu.wpi.first.util.datalog.DataLogRecord;

public abstract class PearascopeProcessor {
    public void run(boolean enableMonitoring, boolean generateRawDump, String[] incomingFilePaths) {
        if (enableMonitoring) {
            // just monitor the input folder for new logs
            Path folder = Paths.get("./input");

            try (WatchService watchService = FileSystems.getDefault().newWatchService()) {
                folder.register(watchService, StandardWatchEventKinds.ENTRY_CREATE);

                System.out.println("Monitoring folder: " + folder.toAbsolutePath());

                while (true) {
                    WatchKey key = watchService.take();

                    for (WatchEvent<?> event : key.pollEvents()) {
                        WatchEvent.Kind<?> kind = event.kind();

                        if (kind == StandardWatchEventKinds.ENTRY_CREATE) {
                            Path filename = (Path) event.context();
                            Path filePath = folder.resolve(filename);
                            if (filePath.toAbsolutePath().toString().toLowerCase().endsWith(".wpilog")) {
                                while (!Utils.isFileCompletelyWritten(filePath)) {
                                    Thread.sleep(500);
                                }
                                processLogs(List.of(filePath.toAbsolutePath().toString()), false);
                            }
                        }
                    }

                    // Reset the key to receive further watch events
                    boolean valid = key.reset();
                    if (!valid) {
                        System.out.println("Watch key no longer valid, exiting...");
                        break;
                    }
                }
            } catch (IOException | InterruptedException e) {
                e.printStackTrace();
            }
        } else {
            List<String> filePaths = Utils.getWPILogFilePaths(incomingFilePaths);
            if (filePaths.size() == 0) {
                System.err.println("Either provide one or more paths to log files on the command line or place files in the 'input' folder.");
                return;
            }

            processLogs(filePaths, generateRawDump);
        }
    }

    protected void processLogs(List<String> filePaths, boolean generateRawDump) {
        for (String logFilePath : filePaths) {
            if (generateRawDump) dumpRawLog(logFilePath);

            String outputFilePath = Utils.getOutputFilePath(logFilePath);
            File file = new File(outputFilePath + ".xlsx");
            if (file.exists()) {
                System.out.println("Output already exists for " + file.getName());
                continue;
            }

            System.out.println("Processing " + logFilePath);

            processLog(logFilePath, outputFilePath);
        }

        // this should not be necessary, but if monitoring is enabled, because WPILib's DataLogReader doesn't properly close and dispose of the file handle, we need to encourage garbage collection to dispose of the handle
        System.gc(); 
        try {
            Thread.sleep(100); // Let GC settle
        } catch (InterruptedException e) {
            // Ignore
        }
    }

    private void processLog(String logFilePath, String outputFilePath) {
        DataLogReader reader;
        try {
            reader = new DataLogReader(logFilePath);
        } catch (IOException ex) {
            System.err.println("ERROR: could not open file: " + ex.getMessage());
            return;
        }
        if (!reader.isValid()) {
            System.err.println("ERROR: not a log file");
            return;
        }

        processLog(reader, outputFilePath);
    }
    
    protected abstract void processLog(DataLogReader reader, String outputFilePath);

    protected abstract int outputEntriesOfInterest(XSSFSheet sheet, int rowIndex, String matchPeriod, double matchTime, DataLogRecord record, DataLogRecord.StartRecordData entry);

    protected abstract void formatOutput(XSSFWorkbook workbook, XSSFSheet sheet);

    protected abstract void dumpRawLog(String logFilePath);
}