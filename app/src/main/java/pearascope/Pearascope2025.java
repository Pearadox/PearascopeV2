package pearascope;

import edu.wpi.first.util.datalog.DataLogReader;
import edu.wpi.first.util.datalog.DataLogRecord;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.InputMismatchException;
import java.util.Map;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;
import org.apache.poi.xssf.usermodel.*;

public class Pearascope2025 extends PearascopeProcessor {
    // assuming for now that these columns might change year-over-year - but if we determine they can be the same every year (like with a standard analysis output format), then we could move this enum to the Excel helper class and remove the ColumnEnum interface
    private enum COLUMN implements ColumnEnum {
        TIMESTAMP,
        MATCHPERIOD,
        MATCHTIME,
        ID,
        ENTRY,
        DATATYPE,
        VALUE,
        VALUERAW,
        PIECE,
        ACTION,
        ACTIONDATA,
        CYCLETIME,
        OUTCOME;
    }

    @Override
    protected void processLog(DataLogReader reader, String outputFilePath) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("data");
        String[] headers = new String[] {"Timestamp", "Period", "M_Time", "ID", "Entry", "Type", "Value", "Value_RAW", "Piece", "Action", "T/Val", "Cycle", "Outcome"};
        int rowIndex = Excel.addOutputHeader(sheet, 0, headers);

        int records = 0;
        long timestamp = -1;
        String matchPeriod = "";
        double matchTime = 0.0;
        boolean auto = false;
        Map<Integer, DataLogRecord.StartRecordData> entries = new HashMap<>();
        try {
            for (DataLogRecord record : reader) {
                // Keeping track of how many records we're processing - this is only used for console output to inform the user.
                if (timestamp != record.getTimestamp()) {
                    records++;
                    timestamp = record.getTimestamp();
                }
                // The wpilog spec allows for other record types, but our logs only seem to carry start and data records.
                if (record.isStart()) {
                    try {
                        DataLogRecord.StartRecordData data = record.getStartData();
                        entries.put(data.entry, data);
                    } catch (InputMismatchException ex) {
                        System.err.println("WARNING: Start(INVALID)");
                    }
                } else {
                    DataLogRecord.StartRecordData entry = entries.get(record.getEntry());
                    if (entry == null) {
                        System.err.println("WARNING: <ID not found: " + record.getEntry() + ">");
                        continue;
                    }

                    if (entry.name.equals("/DriverStation/Autonomous")) auto = record.getBoolean();
                    if (entry.name.equals("/DriverStation/MatchTime")) matchTime = record.getDouble();                        
                    matchPeriod = Utils.updateMatchPeriod(matchPeriod, auto, record, entry);
                    
                    if (matchPeriod.equals("teleop")) {
                        // this ensures we're only analyzing teleop - we could || "auto" to also look at auto
                        rowIndex = outputEntriesOfInterest(sheet, rowIndex, matchPeriod, matchTime, record, entry);
                    } else if (entry.name.equals("/DriverStation/Enabled") && matchPeriod.equals("match end")) {
                        // once we've reached match end, output a final row - all entries after this will get skipped
                        addOutputRow(sheet, ++rowIndex, matchPeriod, matchTime, record, entry, matchPeriod);
                    }
                }
            }
        } catch (IllegalArgumentException ex) {
            System.err.println("WARNING: IllegalArgumentException (might be fine - check the output)");
        }

        addCycleTimes(sheet);
        addIntakeTimes(sheet);
        addStrafingTimes(sheet);
        addIntakeToOuttakeTimes(sheet);
        addAlignTimes(sheet);
        addPieceLabels(sheet);
        addElevatorHoming(sheet);
        addClimbTimes(sheet);

        formatOutput(workbook, sheet);
        int maxRow = sheet.getLastRowNum() + 1;

        addSummaryAnalysis(sheet);

        try (FileOutputStream fileOut = new FileOutputStream(outputFilePath + ".xlsx")) {
            workbook.write(fileOut);
            workbook.close();
            System.out.println("Excel file created successfully: " + outputFilePath + ".xlsx");
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println(records + " records processed [" + maxRow + " rows in output]");
    }

    @Override
    protected int outputEntriesOfInterest(XSSFSheet sheet, int rowIndex, String matchPeriod, double matchTime, DataLogRecord record, DataLogRecord.StartRecordData entry) {
        if(entry.name.equals("/RealOutputs/EE/Has Coral")) {
            if(record.getBoolean()) {
                addOutputRow(sheet, ++rowIndex, matchPeriod, matchTime, record, entry, "intake done");
            }
        }

        // autodrive to target; 90 = right; 270 = left; 180 = algae; 0 = station; -1 = no button pressed
        if(entry.name.equals("/DriverStation/Joystick0/POVs")) {
            long[] vals = record.getIntegerArray();
            if(vals.length > 0) {
                String alignStr = "";
                if(vals[0] == 0) {
                    alignStr = "align station";
                } else if(vals[0] == 90) {
                    alignStr = "align right";
                } else if(vals[0] == 180) {
                    alignStr = "align algae";
                } else if(vals[0] == 270) {
                    alignStr = "align left";
                } else if(vals[0] == -1) {
                    alignStr = "(align release)";
                }
                addOutputRow(sheet, ++rowIndex, matchPeriod, matchTime, record, entry, alignStr, vals[0]);
            }
        }

        if(entry.name.equals("/RealOutputs/Align/Error/IsAligned") || entry.name.equals("/RealOutputs/Align/Error/IsAlignedTest")) {
            String aligned = (record.getBoolean() ? "Aligned" : "Not aligned");
            Row row = addOutputRow(sheet, ++rowIndex, matchPeriod, matchTime, record, entry, String.valueOf(record.getBoolean()));
            // setCellValue(row, COLUMN.ACTION, aligned);
        }

        // 16 = intake, 32 = outtake, 1 = slow toggle
        // (This is actually a bitmask, but unless they press two buttons at once, we can just check the int values for simplicity.  If we do ever use the bitmask: 0 A, 1 B, 2 X, 3 Y, 4 LBump, 5 RBump, 6 Back, 7 Start, 8 LStick, 9 RStick.
        if(entry.name.equals("/DriverStation/Joystick0/ButtonValues")) {
            long button = record.getInteger();
            if (button == 16 || button == 32) {
                addOutputRow(sheet, ++rowIndex, matchPeriod, matchTime, record, entry, (button == 16 ? "intake" : "outtake"), button);
            } else if (button == 1) {
                Row row = addOutputRow(sheet, ++rowIndex, matchPeriod, matchTime, record, entry, "SLOW TOGGLE", button);
                Excel.setCellValue(row, COLUMN.ACTION, "Slow toggle");
            }
        }

        // /2 - strafe left; /3 - strafe right
        if(entry.name.equals("/DriverStation/Joystick0/AxisValues")) {
            float[] axisValues = record.getFloatArray();
            if(axisValues.length >= 4 && (axisValues[2] > 0.0 || axisValues[3] > 0.0)) {
                if (axisValues[2] > 0.0) {
                    addOutputRow(sheet, ++rowIndex, matchPeriod, matchTime, record, entry, "<- strafe", axisValues[2]);
                } else {
                    addOutputRow(sheet, ++rowIndex, matchPeriod, matchTime, record, entry, "strafe ->", axisValues[3]);
                }
            }
        }

        // climb; 180 = deploy; 0 = retract 
        if(entry.name.equals("/DriverStation/Joystick1/POVs")) {
            long[] vals = record.getIntegerArray();
            if(vals.length > 0 && (vals[0] == 0 || vals[0] == 180)) {
                String alignStr = "";
                if(vals[0] == 0) {
                    alignStr = "climb retract";
                } else if(vals[0] == 180) {
                    alignStr = "climb deploy";
                }
                addOutputRow(sheet, ++rowIndex, matchPeriod, matchTime, record, entry, alignStr, vals[0]);
            }
        }

        // 128 = home elevator
        // (this is actually a bitmask, but unless they press two buttons at once, we can just check the int values for simplicity)
        if(entry.name.equals("/DriverStation/Joystick1/ButtonValues")) {
            long button = record.getInteger();
            if (button == 128) {
                addOutputRow(sheet, ++rowIndex, matchPeriod, matchTime, record, entry, "ELEV HOME", button);
            }
        }

        // could monitor for joystick1 axisvalues, but easier to just monitor the offset - this might be an issue if we comment out logging for elevator offset (joystick axisvalues is lower layer)
        if(entry.name.equals("/RealOutputs/Elevator/Offset")) {
            addOutputRow(sheet, ++rowIndex, matchPeriod, matchTime, record, entry, "ELEV OFFSET", record.getDouble());
        }

        if(entry.name.equals("/RealOutputs/Arm/Mode") || entry.name.equals("/RealOutputs/Elevator Mode")) {
            addOutputRow(sheet, ++rowIndex, matchPeriod, matchTime, record, entry, record.getString().toUpperCase());
        }

        return rowIndex;
    }

    private Row addOutputRow(Sheet sheet, int rowIndex, String matchPeriod, double matchTime, DataLogRecord record, DataLogRecord.StartRecordData entry, String valueStr) {
        return addOutputRow(sheet, rowIndex, matchPeriod, matchTime, record, entry, valueStr, Double.MIN_VALUE);
    }

    private Row addOutputRow(Sheet sheet, int rowIndex, String matchPeriod, double matchTime, DataLogRecord record, DataLogRecord.StartRecordData entry, String valueStr, double valueRaw) {
        Row row = sheet.createRow(rowIndex);
        Excel.setCellValue(row, COLUMN.TIMESTAMP, record.getTimestamp() / 1000000.0);
        Excel.setCellValue(row, COLUMN.MATCHPERIOD, matchPeriod);
        Excel.setCellValue(row, COLUMN.MATCHTIME, matchTime);
        Excel.setCellValue(row, COLUMN.MATCHTIME, String.format("%02d:%02d", (int)matchTime / 60, (int)matchTime % 60));
        Excel.setCellValue(row, COLUMN.ID, record.getEntry());
        Excel.setCellValue(row, COLUMN.ENTRY, entry.name);
        Excel.setCellValue(row, COLUMN.DATATYPE, entry.type);
        Excel.setCellValue(row, COLUMN.VALUE, valueStr);
        if (Math.abs(valueRaw) != Double.MIN_VALUE)
        Excel.setCellValue(row, COLUMN.VALUERAW, valueRaw);

        return row; 
    }

    private void addCycleTimes(Sheet sheet) {
        int maxRow = sheet.getLastRowNum() + 1;
        int startRow = 0;
        
        for (int r = 2; r <= maxRow; r++) {
            Row row = sheet.getRow(r - 1);
            if(Excel.getCellString(row, COLUMN.VALUE) == "outtake") {
                if(startRow > 0) {
                    Excel.setCellFormula(row, COLUMN.CYCLETIME, "A" + r + "-A" + startRow);
                }
                startRow = r;
            }
        }
    }

    private void addIntakeTimes(Sheet sheet) {
        int maxRow = sheet.getLastRowNum() + 1;
        int startRow = 0;
        
        for (int r = 2; r <= maxRow; r++) {
            Row row = sheet.getRow(r - 1);
            if(Excel.getCellString(row, COLUMN.VALUE) == "intake" && startRow == 0) {
                startRow = r;
            }
            else if(Excel.getCellString(row, COLUMN.VALUE) == "intake done" && startRow > 0) {
                Excel.setCellValue(row, COLUMN.ACTION, "Time to intake");
                Excel.setCellFormula(row, COLUMN.ACTIONDATA, "A" + r + "-A" + startRow);
                startRow = 0;
            }
        }
    }

    private void addStrafingTimes(Sheet sheet) {
        int maxRow = sheet.getLastRowNum() + 1;
        int startRow = 0;
        int endRow = 0;
        
        for (int r = 2; r <= maxRow; r++) {
            Row row = sheet.getRow(r - 1);
            if(Excel.getCellString(row, COLUMN.VALUE).contains("strafe") && startRow == 0) {
                startRow = r;
            } else if((Excel.getCellString(row, COLUMN.VALUE) == "intake done" || Excel.getCellString(row, COLUMN.VALUE) == "outtake") && startRow > 0 && endRow > startRow) {
                Row row2 = sheet.getRow(endRow - 1);
                Excel.setCellValue(row2, COLUMN.ACTION, "Time spent strafing");
                Excel.setCellFormula(row2, COLUMN.ACTIONDATA, "A" + endRow + "-A" + startRow);
                startRow = 0;
                endRow = 0;
            } else if (Excel.getCellString(row, COLUMN.VALUE).contains("strafe")) {
                endRow = r;
            }
        }
    }

    private void addIntakeToOuttakeTimes(Sheet sheet) {
        int maxRow = sheet.getLastRowNum() + 1;
        int startRow = 0;
        String piece = "";
        String level = "";
        
        for (int r = 2; r <= maxRow; r++) {
            Row row = sheet.getRow(r - 1);
            if (Excel.getCellString(row, COLUMN.VALUE).endsWith("algae")) {
                piece = "ALGAE";
            } else if (Excel.getCellString(row, COLUMN.VALUE).startsWith("align")) {
                piece = "CORAL";
            } else if (Excel.getCellString(row, COLUMN.ENTRY).equals("/RealOutputs/Arm/Mode")) {
                level = Excel.getCellString(row, COLUMN.VALUE);
            }
            if (Excel.getCellString(row, COLUMN.VALUE) == "intake") {
                startRow = 0;
            } else if(Excel.getCellString(row, COLUMN.VALUE) == "intake done" && startRow == 0) {
                startRow = r;
            } else if (Excel.getCellString(row, COLUMN.VALUE) == "outtake" && startRow > 0) {
                Excel.setCellValue(row, COLUMN.ACTION, "In to out");
                Excel.setCellFormula(row, COLUMN.ACTIONDATA, "A" + r + "-A" + startRow);
                String outcome = piece;
                if (piece.equals("CORAL") && level.length() > 0) outcome += " " + level;
                Excel.setCellValue(row, COLUMN.OUTCOME, outcome);
                startRow = 0;
            }
        }
    }

    private void addAlignTimes(Sheet sheet) {
        int maxRow = sheet.getLastRowNum() + 1;
        int startRow = 0;
        int endRow = 0;
        
        for (int r = 2; r <= maxRow; r++) {
            Row row = sheet.getRow(r - 1);
            if(!(Excel.getCellString(row, COLUMN.ENTRY).equals("/RealOutputs/Arm/Mode") || Excel.getCellString(row, COLUMN.ENTRY).equals("/RealOutputs/Elevator Mode"))) {
                if(Excel.getCellString(row, COLUMN.VALUE).startsWith("align") && startRow == 0) {
                    startRow = r;
                }
                else if(Excel.getCellString(row, COLUMN.VALUE).startsWith("align") && endRow > startRow) {
                    endRow = 0;
                } else if (!Excel.getCellString(row, COLUMN.VALUE).contains("align") && startRow > 0 && endRow > startRow) {
                    Row row2 = sheet.getRow(endRow - 1);
                    Excel.setCellValue(row2, COLUMN.ACTION, "Time spent aligning");
                    Excel.setCellFormula(row2, COLUMN.ACTIONDATA, "A" + endRow + "-A" + startRow);
                    startRow = 0;
                    endRow = 0;
                } else if (Excel.getCellString(row, COLUMN.VALUE) == "(align release)") {
                    endRow = r;
                }
            }
        }
    }

    private void addPieceLabels(Sheet sheet) {
        int maxRow = sheet.getLastRowNum() + 1;
        int piece = 1;

        for (int r = 2; r <= maxRow; r++) {
            Row row = sheet.getRow(r - 1);
            if (row.getCell(COLUMN.ACTION.ordinal()) != null && !Excel.getCellString(row, COLUMN.VALUE).equals("SLOW TOGGLE")) {
                Excel.setCellValue(row, COLUMN.PIECE, "Piece " + piece);
                if (Excel.getCellString(row, COLUMN.VALUE) == "outtake") piece++;
            }
        }
    }

    private void addElevatorHoming(Sheet sheet) {
        int maxRow = sheet.getLastRowNum() + 1;
        double offset = 0.0;

        for (int r = 2; r <= maxRow; r++) {
            Row row = sheet.getRow(r - 1);
            if(Excel.getCellString(row, COLUMN.VALUE).equals("ELEV OFFSET")) {
                double valraw = Excel.getCellNumber(row, COLUMN.VALUERAW);
                if(Math.abs(valraw) > 0.0) {
                    offset = Excel.getCellNumber(row, COLUMN.VALUERAW);
                }
            } else if(Excel.getCellString(row, COLUMN.VALUE).equals("ELEV HOME")) {
                Excel.setCellValue(row, COLUMN.ACTION, "Elevator homed/zeroed");
                Excel.setCellValue(row, COLUMN.ACTIONDATA, offset);
            }
        }
    }

    private void addClimbTimes(Sheet sheet) {
        int maxRow = sheet.getLastRowNum() + 1;
        double firstTimestamp = Excel.getCellNumber(sheet.getRow(1), COLUMN.TIMESTAMP);
        int startRow = 0;
        int endRow = 0;
        
        for (int r = 2; r <= maxRow; r++) {
            Row row = sheet.getRow(r - 1);
            if (Excel.getCellNumber(row, COLUMN.TIMESTAMP) > firstTimestamp + 60.0) {
                if (Excel.getCellString(row, COLUMN.VALUE).startsWith("climb") && startRow == 0) {
                    startRow = r;
                } else if ((!Excel.getCellString(row, COLUMN.VALUE).startsWith("climb") && endRow > startRow)) {
                    Row row2 = sheet.getRow(endRow - 1);
                    Excel.setCellValue(row2, COLUMN.ACTION, "Time spent climbing");
                    Excel.setCellFormula(row2, COLUMN.ACTIONDATA, "A" + endRow + "-A" + startRow);
                    startRow = 0;
                    endRow = 0;
                } else if (r == maxRow && startRow > 0) {
                    Excel.setCellValue(row, COLUMN.ACTION, "Time spent climbing");
                    Excel.setCellFormula(row, COLUMN.ACTIONDATA, "A" + r + "-A" + startRow);
                    startRow = 0;
                    endRow = 0;
                } else if (Excel.getCellString(row, COLUMN.VALUE).startsWith("climb") && r < maxRow) {
                    endRow = r;
                }
            }
        }
    }

    @Override
    protected void formatOutput(XSSFWorkbook workbook, XSSFSheet sheet) {
        byte[] clr_lightred = new byte[] {(byte)255, (byte)199, (byte)206};
        byte[] clr_darkred = new byte[] {(byte)156, (byte)0, (byte)6};
        byte[] clr_lightyellow = new byte[] {(byte)255, (byte)235, (byte)156};
        byte[] clr_darkyellow = new byte[] {(byte)156, (byte)87, (byte)0};
        byte[] clr_lightgreen = new byte[] {(byte)198, (byte)239, (byte)206};
        byte[] clr_darkgreen = new byte[] {(byte)0, (byte)97, (byte)0};
        byte[] clr_lightblue = new byte[] {(byte)166, (byte)201, (byte)236};
        byte[] clr_darkblue = new byte[] {(byte)21, (byte)61, (byte)100};

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        ConditionalFormattingRule strafRule = Excel.highlightTextContaining(sheetCF, "straf", clr_lightred, clr_darkred);  //purposefully "straf" to catch "strafe" and "strafing"
        ConditionalFormattingRule intakeRule = Excel.highlightTextContaining(sheetCF, "intake", clr_lightyellow, clr_darkyellow);
        ConditionalFormattingRule alignRule = Excel.highlightTextContaining(sheetCF, "align", clr_lightgreen, clr_darkgreen);
        ConditionalFormattingRule outtakeRule = Excel.highlightTextContaining(sheetCF, "out", clr_lightblue, clr_darkblue);

        int maxRow = sheet.getLastRowNum() + 1;
        CellRangeAddress[] regions = {
            CellRangeAddress.valueOf("G1:G" + maxRow),
            CellRangeAddress.valueOf("J1:J" + maxRow)
        };
        sheetCF.addConditionalFormatting(regions, new ConditionalFormattingRule[] { outtakeRule, strafRule, intakeRule, alignRule });

        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        for (int i = 0; i < maxRow; i++) {
            Cell cell = sheet.getRow(i).getCell(COLUMN.VALUE.ordinal());
            if (cell != null) cell.setCellStyle(style);
        }

        CellStyle cellStyle = workbook.createCellStyle();
        DataFormat dataFormat = workbook.createDataFormat();
        cellStyle.setDataFormat(dataFormat.getFormat("#,##0.00"));
        for (int i = 1; i < maxRow; i++) {
            Cell time = sheet.getRow(i).getCell(COLUMN.ACTIONDATA.ordinal());
            if (time != null) time.setCellStyle(cellStyle);
            Cell cycle = sheet.getRow(i).getCell(COLUMN.CYCLETIME.ordinal());
            if (cycle != null) cycle.setCellStyle(cellStyle);
        }

        sheet.setColumnWidth(COLUMN.TIMESTAMP.ordinal(), 11 * 256);
        sheet.setColumnWidth(COLUMN.MATCHPERIOD.ordinal(), 9 * 256);
        sheet.setColumnWidth(COLUMN.MATCHTIME.ordinal(), 7 * 256);
        sheet.setColumnWidth(COLUMN.ID.ordinal(), 0);
        sheet.setColumnWidth(COLUMN.ENTRY.ordinal(), 31 * 256);
        sheet.setColumnWidth(COLUMN.DATATYPE.ordinal(), 0);
        sheet.setColumnWidth(COLUMN.VALUE.ordinal(), 13 * 256);
        sheet.setColumnWidth(COLUMN.VALUERAW.ordinal(), 12 * 256);
        sheet.setColumnWidth(COLUMN.PIECE.ordinal(), 8 * 256);
        sheet.setColumnWidth(COLUMN.ACTION.ordinal(), 17 * 256);
        sheet.setColumnWidth(COLUMN.ACTIONDATA.ordinal(), 7 * 256);
        sheet.setColumnWidth(COLUMN.CYCLETIME.ordinal(), 7 * 256);
        sheet.setColumnWidth(COLUMN.OUTCOME.ordinal(), 20 * 256);

        sheet.createFreezePane(0, 1);
        
        Excel.convertToTable(workbook, sheet);

        Excel.filterValues(sheet, COLUMN.ACTION, "Time to intake", "Time spent aligning", "In to out", "Time spent strafing", "Time spent climbing", "Elevator homed/zeroed", "Slow toggle");
    }

    private void addSummaryAnalysis(XSSFSheet sheet) {
        int maxRow = sheet.getLastRowNum();

        Row row = sheet.createRow(maxRow + 2);
        row.createCell(0).setCellValue("Avg Cycle Time");
        row.createCell(2).setCellFormula("AVERAGE(L:L)");
        
        row = sheet.createRow(maxRow + 3);
        row.createCell(0).setCellValue("Avg Time Spent Intaking");
        row.createCell(2).setCellFormula("AVERAGEIF(J:J, \"Time to intake\", K:K)");

        row = sheet.createRow(maxRow + 4);
        row.createCell(0).setCellValue("Avg Time Spent Strafing");
        row.createCell(2).setCellFormula("AVERAGEIF(J:J, \"Time spent strafing\", K:K)");

        row = sheet.createRow(maxRow + 5);
        row.createCell(0).setCellValue("Avg Time Spent Aligning");
        row.createCell(2).setCellFormula("AVERAGEIF(J:J, \"Time spent aligning\", K:K)");

        row = sheet.createRow(maxRow + 6);
        row.createCell(0).setCellValue("Game pieces handled");
        row.createCell(2).setCellFormula("SUMPRODUCT((I2:I" + maxRow + "<>\"\")/COUNTIF(I2:I" + maxRow + ", I2:I" + maxRow + "&\"\"))");
    }

    @Override
    protected void dumpRawLog(String logFilePath) {
        System.out.println("Processing log (RAW) " + logFilePath);
            
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

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("data");
        String[] headers = new String[] {"Timestamp", "Period", "M_Time", "ID", "Entry", "Type", "Value_RAW"};
        int rowIndex = Excel.addOutputHeader(sheet, 0, headers);

        int records = 0;
        long timestamp = -1;
        String matchPeriod = "";
        double matchTime = 0.0;
        boolean auto = false;
        Map<Integer, DataLogRecord.StartRecordData> entries = new HashMap<>();
        try {
            for (DataLogRecord record : reader) {
                if (timestamp != record.getTimestamp()) {
                    records++;
                    timestamp = record.getTimestamp();
                }
                if (record.isStart()) {
                    try {
                        DataLogRecord.StartRecordData data = record.getStartData();
                        entries.put(data.entry, data);
                    } catch (InputMismatchException ex) {
                        System.err.println("WARNING: Start(INVALID)");
                    }
                } else {
                    DataLogRecord.StartRecordData entry = entries.get(record.getEntry());
                    if (entry == null) {
                        System.err.println("WARNING: <ID not found: " + record.getEntry() + ">");
                        continue;
                    }

                    if(entry.name.equals("/DriverStation/Autonomous")) auto = record.getBoolean();
                    matchPeriod = Utils.updateMatchPeriod(matchPeriod, auto, record, entry);

                    if (entry.name.equals("/DriverStation/MatchTime")) matchTime = record.getDouble();                        

                    // The logic below ensures we only dump log for the match itself, not junk before or after while the bot is still powered on.
                    if(matchPeriod.equals("auto") || matchPeriod.equals("disabled") || matchPeriod.equals("teleop") || entry.name.equals("/DriverStation/Enabled")) {
                        // This does not handle custom structs, which currently includes ChassisSpeeds, Pose2d, Rotation2d, SwerveModulePosition, SwerveModuleState, Transform2d, & Translation2d.  Processing these takes special handling that I haven't yet sussed out.
                        switch (entry.type) {
                            case "float" -> addOutputRow(sheet, ++rowIndex, matchPeriod, matchTime, record, entry, String.valueOf(record.getFloat())); 
                            case "double" -> addOutputRow(sheet, ++rowIndex, matchPeriod, matchTime, record, entry, String.valueOf(record.getDouble()));
                            case "int64" -> addOutputRow(sheet, ++rowIndex, matchPeriod, matchTime, record, entry, String.valueOf(record.getInteger()));
                            case "string", "json" -> addOutputRow(sheet, ++rowIndex, matchPeriod, matchTime, record, entry, record.getString());
                            case "boolean" -> addOutputRow(sheet, ++rowIndex, matchPeriod, matchTime, record, entry, String.valueOf(record.getBoolean()));
                            case "float[]" -> {
                                String values = Arrays.toString(record.getFloatArray()).replaceAll("[\\[\\] ]", "");
                                addOutputRow(sheet, ++rowIndex, matchPeriod, matchTime, record, entry, values);
                            }
                            case "double[]" -> {
                                String values = Arrays.toString(record.getDoubleArray()).replaceAll("[\\[\\] ]", "");
                                addOutputRow(sheet, ++rowIndex, matchPeriod, matchTime, record, entry, values);
                            }
                            case "int64[]" -> {
                                String values = Arrays.toString(record.getIntegerArray()).replaceAll("[\\[\\] ]", "");
                                addOutputRow(sheet, ++rowIndex, matchPeriod, matchTime, record, entry, values);
                            }
                            case "string[]" -> addOutputRow(sheet, ++rowIndex, matchPeriod, matchTime, record, entry, String.join(",", record.getStringArray()));
                            default -> { }
                        }
                    }
                }
            }
        } catch (IllegalArgumentException ex) {
            System.err.println("WARNING: IllegalArgumentException (might be fine - check the output)");
        }

        String outputFilePath = Utils.getOutputFilePath(logFilePath);

        int maxRow = sheet.getLastRowNum() + 1;

        try (FileOutputStream fileOut = new FileOutputStream(outputFilePath + ".RAW.xlsx")) {
            workbook.write(fileOut);
            workbook.close();
            System.out.println("Excel file (RAW) created successfully: " + outputFilePath + ".RAW.xlsx");
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println(records + " records processed (RAW) [" + maxRow+ " rows in output]");
    }
}
