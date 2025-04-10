package pearascope;

import java.io.File;
import java.io.IOException;
import java.nio.channels.FileChannel;
import java.nio.channels.FileLock;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import edu.wpi.first.util.datalog.DataLogRecord;

public class Utils {
    public static String updateMatchPeriod(String matchPeriod, boolean auto, DataLogRecord record, DataLogRecord.StartRecordData entry) {
        // Once the robot is connected to the FMS, autonomous will be TRUE and enabled will be FALSE.  So, once we see enabled go TRUE while auto is TRUE, we know the Auto period has begun.  After Auto period, there will be a momentary (~2 seconds) Disabled period.  When enabled goes TRUE again, Teleop period has begun - autonomous will be false, but we don't need to check that as we're watching for the transition from Disabled period.
        if(entry.name.equals("/DriverStation/Enabled") && record.getBoolean() && auto) {
            matchPeriod = "auto";
        } else if (entry.name.equals("/DriverStation/Enabled") && !record.getBoolean() && matchPeriod.equals("auto")) {
            matchPeriod = "disabled";
        } else if (entry.name.equals("/DriverStation/Enabled") && record.getBoolean() && matchPeriod.equals("disabled")) {
            matchPeriod = "teleop";
        } else if (entry.name.equals("/DriverStation/Enabled") && !record.getBoolean() && matchPeriod.equals("teleop")) {
            matchPeriod = "match end";
        }
        return matchPeriod;
    }

    public static boolean isFileCompletelyWritten(Path path) {
        try (FileChannel channel = FileChannel.open(path, StandardOpenOption.READ)) {
            FileLock lock = channel.tryLock(0, Long.MAX_VALUE, true);
            if (lock != null) {
                lock.release();
                channel.close();
                return true;
            }
            channel.close();
        } catch (IOException e) {
            // File is likely still being written
        }
        return false;
    }

    public static String getOutputFilePath(String logFilePath) {
        Pattern pattern = Pattern.compile("akit_(\\d{2})-.*_(\\w+_\\w+)\\.wpilog");
        Matcher matcher = pattern.matcher(logFilePath);
        String matchPrefix = "";
        if (matcher.find()) {
            String year = matcher.group(1);
            String match = matcher.group(2);
            matchPrefix = "20" + year + match;
        }

        File file = new File(logFilePath);
        String folderPath = file.getParent();

        return folderPath + "\\" + matchPrefix;
    }

    public static List<String> getWPILogFilePaths(String[] args) {
        List<String> filePaths = new ArrayList<>();
        if(args.length > 0) {
            for (String filePath : args) {
                if (filePath.toLowerCase().endsWith(".wpilog"))
                    filePaths.add(filePath);
            }
        } else {
            // No command line input - process everything in the ./input/ folder
            File inputFolder = new File(System.getProperty("user.dir"), "input");
            if (inputFolder.exists() && inputFolder.isDirectory()) {
                File[] files = inputFolder.listFiles();
                if (files != null) {
                    for (File file : files) {
                        if (file.isFile() && file.getPath().toLowerCase().endsWith(".wpilog")) {
                            filePaths.add(file.getAbsolutePath());
                        }
                    }
                }
            }
        }
        return filePaths;
    }
}
