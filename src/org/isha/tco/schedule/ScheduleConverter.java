/*
 * Copyright (c) 2009 Isha Foundation. All rights reserved.
 */

package org.isha.tco.schedule;

import java.io.File;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.MissingResourceException;
import java.util.ResourceBundle;
import java.util.Set;
import java.util.TreeMap;

import jxl.Cell;
import jxl.Range;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.DateFormat;
import jxl.write.DateTime;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/**
 * 
 * @author psriniv
 *
 */
public class ScheduleConverter {

    private static final int EXCEL_START_COL = 0;
    private static final int MONTH_YEAR_ROW = 1;
    private static final int DATE_OF_MONTH_ROW = MONTH_YEAR_ROW + 1;
    private static final int TEACHER_START_COL = EXCEL_START_COL + 1;
    private static final int TEACHER_START_ROW = DATE_OF_MONTH_ROW + 1;

    /** top-left, bottom-right to the contents in the merged cells */
    private static final Map<CellInfo, Map<CellInfo, String>> MERGED_CELLS_MAP = new HashMap<CellInfo, Map<CellInfo,String>>(); 

    static interface ConfigKey {
        static final String DEBUG = "debug";
        static final String OUTPUT_DIRECTORY = "output_directory";
        static final String PLACE_OWNER_WORKBOOK_FILENAME = "place_owner_workbook_filename";
        static final String ACTIVITIES_FOR_GROUPING_TEACHERS = "activities_for_grouping_teachers";
    }

    static interface ReportFilterType {
        static final String ALL = "All";
        static final String TEACHER = "Teacher";
        static final String SECTOR_COORDINATOR = "SectorCoordinator";
    }

    static interface OutputSuffix {
        static final String CONSOLIDATED = "ConsolidatedReport.xls";
        static final String PER_TEACHER_FILE = ".xls";
        static final String PER_COORD_FILE = ".xls";
        static final String PER_TEACHER_DIR = File.separator + "teachers" + File.separator;
        static final String PER_COORD_DIR = File.separator + "coords" + File.separator;
    }

    private static boolean isDebug = false;

    /**
     * @param args
     * @throws Exception 
     */
    public static void main(String[] args) throws Exception {
        ResourceBundle props = ResourceBundle.getBundle("schedule");

        if(args.length == 0 || "--help".equals(args[0])) {
            usage();
            return;
        }

        try {
            String debugMode = props.getString(ConfigKey.DEBUG);
            isDebug = Boolean.valueOf(debugMode).booleanValue();
        }
        catch(MissingResourceException mre) {
            //ignore this
            System.out.println("Warn: debug config key not found in properties.");
        }
        File inputFile = new File(args[0]);
        Workbook workbook = Workbook.getWorkbook(inputFile);
        Sheet sheet = workbook.getSheet("Chart");

        fillUpMergedCells(sheet);

        DateHelper dh = new DateHelper();
        //Process the months
        dh.processMonths(sheet, MONTH_YEAR_ROW);
        //Process the dates
        dh.processDates(sheet, DATE_OF_MONTH_ROW);

        PlaceOwnerHelper poh = new PlaceOwnerHelper(props);

        // Process the teacher schedule information and store it in output sheet
        String outputFolder = "";
        try {
            outputFolder = props.getString(ConfigKey.OUTPUT_DIRECTORY);
        }
        catch(MissingResourceException mre) {
            System.out.println("Warn: " + ConfigKey.OUTPUT_DIRECTORY + " key not configured in properties. Defaulting output to current directory");
        }
        String prefix = inputFile.getName().substring(0, inputFile.getName().lastIndexOf('.'));
        final String outputFilename = outputFolder + File.separator + prefix + OutputSuffix.CONSOLIDATED;

        makeOutputDirectories(outputFolder);

        String scheduleStartDate = "";
        String scheduleEndDate = "";
        if(args.length > 1) {
            scheduleStartDate = args[1];
        }
        if(args.length > 2) {
            scheduleEndDate = args[2];
        }
        ScheduleHelper sh = new ScheduleHelper(props, dh, poh);
        sh.process(outputFilename, sheet, TEACHER_START_ROW, scheduleStartDate, scheduleEndDate);

        workbook.close();
    }

    private static void makeOutputDirectories(String outputFolder) {
        new File(outputFolder + OutputSuffix.PER_COORD_DIR).mkdirs();
        new File(outputFolder + OutputSuffix.PER_TEACHER_DIR).mkdirs();
    }

    private static void usage() {
        System.out.println("create_schedule.bat <input worksheet name> [<schedule-start-date> [<schedule-end-date]]");
        System.out.println("schedule-start-date and schedule-end-date are expected to be in this format: dd/MMM/YYYY");
    }

    private static void fillUpMergedCells(Sheet sheet) {
        Range[] mergedCells = sheet.getMergedCells();
        for(Range range : mergedCells) {
            Cell topLeft = range.getTopLeft();
            Cell bottomRight = range.getBottomRight();

            CellInfo topLeftInfo = new CellInfo(topLeft);
            CellInfo bottomRightInfo = new CellInfo(bottomRight);

            Map<CellInfo, String> valueMap = new HashMap<CellInfo, String>();
            valueMap.put(bottomRightInfo, topLeft.getContents().trim());
            MERGED_CELLS_MAP.put(topLeftInfo, valueMap);
        }

        if(isDebug) System.out.println(MERGED_CELLS_MAP);
    }

    private static class CellInfo {
        private int col;
        private int row;

        public CellInfo(int col, int row) {
            this.col = col;
            this.row = row;
        }

        public CellInfo(Cell cell) {
            this(cell.getColumn(), cell.getRow());
        }

        /* (non-Javadoc)
         * @see java.lang.Object#hashCode()
         */
        @Override
        public int hashCode() {
            final int prime = 31;
            int result = 1;
            result = prime * result + col;
            result = prime * result + row;
            return result;
        }

        /* (non-Javadoc)
         * @see java.lang.Object#equals(java.lang.Object)
         */
        @Override
        public boolean equals(Object obj) {
            if(this == obj)
                return true;
            if(obj == null)
                return false;
            if(getClass() != obj.getClass())
                return false;
            CellInfo other = (CellInfo) obj;
            if(col != other.col)
                return false;
            if(row != other.row)
                return false;
            return true;
        }

        public String toString() {
            return "[" + col + ", " + row + "]";
        }
    }

    private static class DateHelper {
        private Map<Integer, String> dateMap = new HashMap<Integer, String>();
        private Map<String, List<Integer>> monthMap = new HashMap<String, List<Integer>>();

        public DateHelper() {
            //Dummy constructor
        }

        /**
         * For each column in the given row, create a map from the column index to its contents
         */
        public void processDates(final Sheet sheet, final int row) {
            int columns = sheet.getColumns();
            // Start from 1 since the first column contains teacher names.
            for(int i = TEACHER_START_COL + 1; i < columns; i++) {
                Cell dateCell = sheet.getCell(i, row);
                String dateOfMonth = dateCell.getContents().trim();
                dateMap.put(i, dateOfMonth);
            }

            if(isDebug) System.out.println("date map: " + dateMap);
        }

        /**
         * Create a map from month name to start and end column index of that month.
         */
        public void processMonths(final Sheet sheet, final int row) {
            int columnCount = sheet.getColumns();
            //Start from the second column. First column contains teacher names.
            String prevMonth = "";
            for(int i = TEACHER_START_COL + 1; i < columnCount; i++) {
                Cell monthCell = sheet.getCell(i, row);
                String month = monthCell.getContents().trim();
                if(!"".equals(month)) {
                    //Start of a new month

                    //Store the end col for prev month
                    setEndColumn(prevMonth, i - 1);

                    //Store the start col for this month
                    List<Integer> startEndCol = new ArrayList<Integer>();
                    startEndCol.add(i);
                    monthMap.put(month, startEndCol);

                    //Change the prev month to the new month.
                    prevMonth = month;
                }
            }

            //Set the end column for the last month
            setEndColumn(prevMonth, columnCount - 1);
            if(isDebug) System.out.println("month map: " + monthMap);
        }

        private void setEndColumn(String month, int endColumn) {
            if(!"".equals(month)) {
                List<Integer> startEndCol = monthMap.get(month);
                startEndCol.add(endColumn);
                monthMap.put(month, startEndCol);
            }
        }

        /**
         * Returns a string of the form &lt;date-of-month>/&lt;month>/&lt;year> for the given column.
         * <p>
         * Example:
         * 15/May/08
         */
        public String getDate(int column) {
            String dateOfMonth = dateMap.get(column);
            String monthYear = "";
            for(String month : monthMap.keySet()) {
                List<Integer> startEndColumns = monthMap.get(month);
                if(startEndColumns.get(0) <= column && column <= startEndColumns.get(1)) {
                    monthYear = month;
                    break;
                }
            }
            if("".equals(monthYear)) {
                throw new IllegalArgumentException("Unable to find the date for the column: " + column);
            }
            //Get month and year from string like May-08
            String[] values = monthYear.split("-");
            if(values.length != 2) {
                System.out.println("Illegal month-year format: " + monthYear + ". Should be in MMM-YY");
            }
            return dateOfMonth + "/" + values[0].trim() + "/" + values[1].trim();
        }
    }

    private static class PlaceOwnerHelper {
        private static final Map<String, String> placeOwnerMap = new HashMap<String, String>();
        private static final int PLACE_OWNER_START_ROW = 2;
        private static final int PLACE_OWNER_START_COL = 1;

        public PlaceOwnerHelper(ResourceBundle props) {
            try {
                String placeOwnerFilename = props.getString(ConfigKey.PLACE_OWNER_WORKBOOK_FILENAME);
                if(placeOwnerFilename != "") {
                    Workbook pohWorkbook = Workbook.getWorkbook(new File(placeOwnerFilename));
                    Sheet placeOwnerSheet = pohWorkbook.getSheet("Place Owner Table");
                    if(placeOwnerSheet != null) {
                        this.processPlaceOwner(placeOwnerSheet);
                    }
                    else {
                        System.out.println("Warning: Unable to find Place Owner Table sheet. Owners will not be filled.");
                    }
                    pohWorkbook.close();
                }
            }
            catch(MissingResourceException mre) {
                //Ignore this.
                System.out.println("Warn: No property found with ConfigKey: " + ConfigKey.PLACE_OWNER_WORKBOOK_FILENAME);
            } catch(BiffException e) {
                e.printStackTrace();
            } catch(IOException e) {
                e.printStackTrace();
            }
        }

        private void processPlaceOwner(final Sheet placeOwnerSheet) {
            final int MAX_ROW = placeOwnerSheet.getRows();
            for(int row = PLACE_OWNER_START_ROW; row < MAX_ROW; row++) {
                Cell placeCell = placeOwnerSheet.getCell(PLACE_OWNER_START_COL, row);
                Cell ownerCell = placeOwnerSheet.getCell(PLACE_OWNER_START_COL + 1, row);

                placeOwnerMap.put(placeCell.getContents().trim().toLowerCase(), ownerCell.getContents().trim());
            }

            if(isDebug) System.out.println("Place owner map: " + placeOwnerMap);
        }

        public String getOwner(String place) {
            return placeOwnerMap.get(place.toLowerCase());
        }
    }

    private static class ScheduleHelper {
        private static final Set<String> ACTIVITIES_MARKED_FOR_GROUPING = new HashSet<String>();

        private static final String PLACE_ACTIVITY_SEPARATOR = "-";
        private static final String CENTER_SECTOR_SEPARATOR = "/";
        private static final int EXCEL_OUTPUT_START_ROW = 0;
        private static final int EXCEL_OUTPUT_START_COL = 0;
        private DateHelper dateHelper;
        private PlaceOwnerHelper placeOwnerHelper;
        private int maxNumberOfPersons;
        private final Map<String, Set<String>> teacherToLargestGroupMap = new HashMap<String, Set<String>>();

        public ScheduleHelper(ResourceBundle props, DateHelper dh, PlaceOwnerHelper poh) {
            dateHelper = dh;
            placeOwnerHelper = poh;
            try {
                String activitiesForGrouping = props.getString(ConfigKey.ACTIVITIES_FOR_GROUPING_TEACHERS);
                String[] activities = activitiesForGrouping.split(",");
                for(String activity : activities) {
                    ACTIVITIES_MARKED_FOR_GROUPING.add(activity.toLowerCase().trim());
                }
            }
            catch(MissingResourceException mre) {
                mre.printStackTrace();
            }
            System.out.println("Activities marked for grouping set to: " + ACTIVITIES_MARKED_FOR_GROUPING);
        }

        final SimpleDateFormat scheduleDateFormat = new SimpleDateFormat("dd/MMM/yy");
        final Comparator<String> scheduleDateComparator = new Comparator<String>() {
            public int compare(String startDate1, String startDate2) {
                if(startDate1.equals(startDate2))
                    return 0;
                Date start1 = null;
                try {
                    start1 = scheduleDateFormat.parse(startDate1);
                } catch(ParseException e) {
                    e.printStackTrace();
                }
                Date start2 = null;
                try {
                    start2 = scheduleDateFormat.parse(startDate2);
                } catch(ParseException e) {
                    e.printStackTrace();
                }
                
                return start1 == null ? 0 : start1.compareTo(start2);
            }
        };

        public void process(final String outputFilename, final Sheet inputSheet, final int startRow, 
                final String scheduleStartDate, final String scheduleEndDate) 
        throws Exception {

            Map<String, Map<String, Map<String, Map<String, List<String>>>>> startEndPlaceActivityPersonMap = 
                new TreeMap<String, Map<String,Map<String, Map<String, List<String>>>>>(scheduleDateComparator);

            Cell generateOrSkipCell = inputSheet.getCell(EXCEL_START_COL, DATE_OF_MONTH_ROW);
            String generateOrSkip = generateOrSkipCell.getContents().trim();
            boolean skipMarked = "".equals(generateOrSkip) || "skip".equalsIgnoreCase(generateOrSkip);

            System.out.println("Processing mode: " + (skipMarked ? "Skip marked" : "Generate marked"));

            List<String> teachers = new ArrayList<String>();
            List<String> markedTeachers = new ArrayList<String>();
            int endRow = inputSheet.getRows();
            for(int row = startRow; row < endRow; row++) {
                Cell teacherCell = inputSheet.getCell(TEACHER_START_COL, row);
                String teacher = teacherCell.getContents().trim();

                if("".equals(teacher)) {
                    //Nothing to do. Find the next teacher.
                    continue;
                }
                Cell markCell = inputSheet.getCell(EXCEL_START_COL, row);
                String mark = markCell.getContents().toLowerCase().trim();
                if(mark.contains("x")) {
                    markedTeachers.add(teacher);
                }

                teachers.add(teacher);
                //Fill up the map of place-activity to start, end dates for this teacher.
                //Schedules for teacher start from the second row (hence row + 1)
                fillUpPlaceActivityMap(startEndPlaceActivityPersonMap, teacher, inputSheet, 
                        row + 1, scheduleStartDate, scheduleEndDate);
            }

            final File outputFile = new File(outputFilename);
            WritableWorkbook outputWorkbook = Workbook.createWorkbook(outputFile);
            WritableSheet outputSheet = outputWorkbook.createSheet("Output", 0);

            Set<String> sectorCoordinators = 
                writeToExcel(outputSheet, startEndPlaceActivityPersonMap, ReportFilterType.ALL, "");

            if(isDebug) System.out.println("Consolidated schedule: " + startEndPlaceActivityPersonMap);

            outputWorkbook.write();
            outputWorkbook.close();

            //Write a per-teacher sheet.
            writePerTeacherReport(outputFile, teachers, markedTeachers, skipMarked, startEndPlaceActivityPersonMap);

            //Write a per co-oridnator sheet
            writePerCoordinatorReport(outputFile, sectorCoordinators, startEndPlaceActivityPersonMap);
        }

        private void writePerCoordinatorReport(
                final File outputFile, 
                final Set<String> sectorCoordinators,
                final Map<String, Map<String, Map<String, Map<String, List<String>>>>> startEndPlaceActivityPersonMap)
        throws Exception {

            final File parentFolder = outputFile.getParentFile();
            final String folder = parentFolder == null ? "." : parentFolder.getAbsolutePath();
            final String fileName = outputFile.getName();
            final int extensionStartIndex = fileName.lastIndexOf(OutputSuffix.CONSOLIDATED);
            final String prefix = fileName.substring(0, extensionStartIndex);

            for(String coordinator : sectorCoordinators) {
                dateFormat = new WritableCellFormat(new DateFormat("dd-MMM-yyyy"));
                
                final String perCoordinatorOutput = 
                    folder + OutputSuffix.PER_COORD_DIR + prefix + "-" + coordinator + OutputSuffix.PER_COORD_FILE;

                System.out.println("Writting schedule for coordinator: " + coordinator + " to file: " + perCoordinatorOutput);

                WritableWorkbook outputWorkbookPerTeacher = Workbook.createWorkbook(new File(perCoordinatorOutput));
                WritableSheet perTeacherOutputSheet = outputWorkbookPerTeacher.createSheet("Output", 0);

                writeToExcel(perTeacherOutputSheet, startEndPlaceActivityPersonMap, ReportFilterType.SECTOR_COORDINATOR, coordinator);
 
                outputWorkbookPerTeacher.write();
                outputWorkbookPerTeacher.close();
            }
        }

        private void writePerTeacherReport(final File outputFile, 
                final List<String> teachers,
                final List<String> markedTeachers,
                final boolean skipMarked,
                final Map<String, Map<String, Map<String, Map<String, List<String>>>>> startEndPlaceActivityPersonMap)
        throws Exception {

            final File parentFolder = outputFile.getParentFile();
            final String folder = parentFolder == null ? "." : parentFolder.getAbsolutePath();
            final String fileName = outputFile.getName();
            final int extensionStartIndex = fileName.lastIndexOf(OutputSuffix.CONSOLIDATED);
            final String prefix = fileName.substring(0, extensionStartIndex);

            final int maxGroupSize = maxNumberOfPersons;
            for(final String teacher : teachers) {
                dateFormat = new WritableCellFormat(new DateFormat("dd-MMM-yyyy"));

                if(!skipMarked) {
                    //Generate for marked
                    if(!markedTeachers.contains(teacher)) {
                        if(isDebug)
                            System.out.println("Skipping teacher: " + teacher + " as its not marked for generation");
                        continue;
                    }
                }
                else {
                    //Skip those marked
                    if(markedTeachers.contains(teacher)) {
                        if(isDebug)
                            System.out.println("Skipping teacher: " + teacher + " as its marked for skip");
                        continue;
                    }
                }
                Set<String> largestGroup = teacherToLargestGroupMap.get(teacher);
                if(largestGroup == null || largestGroup.size() == 0) {
                    if(isDebug) 
                        System.out.println("Skipping teacher as there is no schedule for this person.");
                    continue;
                }
                maxNumberOfPersons = largestGroup.size(); 

                final String perTeacherOutput = 
                    folder + OutputSuffix.PER_TEACHER_DIR + prefix + "-" + teacher + OutputSuffix.PER_TEACHER_FILE;
                System.out.println("Writting schedule for teacher: " + teacher + " to file: " + perTeacherOutput);

                WritableWorkbook outputWorkbookPerTeacher = Workbook.createWorkbook(new File(perTeacherOutput));
                WritableSheet perTeacherOutputSheet = outputWorkbookPerTeacher.createSheet("Output", 0);

                writeToExcel(perTeacherOutputSheet, startEndPlaceActivityPersonMap, ReportFilterType.TEACHER, teacher);
 
                outputWorkbookPerTeacher.write();
                outputWorkbookPerTeacher.close();
            }
            //Restore it back.
            maxNumberOfPersons = maxGroupSize;
        }
 
        private Set<String> writeToExcel(
                final WritableSheet output,
                final Map<String, Map<String, Map<String, Map<String, List<String>>>>> startEndPlaceActivityPersonMap,
                final String type,
                final String filter) throws Exception {

            Set<String> sectorCoordinators = new HashSet<String>();

            writeHeaderToExcel(output, type);
            int rowPos = EXCEL_OUTPUT_START_ROW + 1;
            int slNo = 1;
            for(String startDate : startEndPlaceActivityPersonMap.keySet()) {
                Map<String, Map<String, Map<String, List<String>>>> endPlaceActivityPersonMap = 
                    startEndPlaceActivityPersonMap.get(startDate);

                for(String endDate : endPlaceActivityPersonMap.keySet()) {
                    Map<String, Map<String, List<String>>> placeActivityPersonMap = 
                        endPlaceActivityPersonMap.get(endDate);

                    for(String place : placeActivityPersonMap.keySet()) {
                        Map<String, List<String>> activityPersonMap = placeActivityPersonMap.get(place);

                        for(String activity : activityPersonMap.keySet()) {
                            List<String> persons = activityPersonMap.get(activity);

                            String center = place;
                            String sector = place;
                            if(place.contains(CENTER_SECTOR_SEPARATOR)) {
                                String[] values = place.split(CENTER_SECTOR_SEPARATOR);
                                center = values[0].trim();
                                sector = values[1].trim();
                                place = center + " " + CENTER_SECTOR_SEPARATOR + " " + sector;
                            }
                            String owner = placeOwnerHelper.getOwner(sector);
                            if(owner == null || "".equals(owner)) {
                                owner = placeOwnerHelper.getOwner(center);
                            }
                            if(!"".equals(owner) && owner != null) {
                                sectorCoordinators.add(owner);
                            }

                            if(ReportFilterType.ALL.equals(type) ||
                                    (ReportFilterType.TEACHER.equals(type) && persons.contains(filter)) ||
                                    (ReportFilterType.SECTOR_COORDINATOR.equals(type) && filter.equals(owner))) {
                                if(ReportFilterType.TEACHER.equals(type)) {
                                    //On per-teacher report, no need for owner.
                                    owner = "";
                                }
                                writeOneRowToExcel(output, rowPos++, String.valueOf(slNo++), 
                                        startDate, endDate, place, activity, persons, owner);
                            }
                        }
                    }
                }
            }

            return sectorCoordinators;
        }

        private void writeHeaderToExcel(WritableSheet output, final String type) throws Exception {
            List<String> persons = new ArrayList<String>();
            for(int i = 1; i <= maxNumberOfPersons; i++) {
                persons.add("Teacher " + i);
            }
            //No need for sector-coordinator if its per teacher report.
            String sectorCoordinatorTitle = ReportFilterType.TEACHER.equals(type) ? "" : "Sector-Coordinator";
            //Sl.No From    To  Center  Nature of Activity  Teacher1  Teacher2  Teacher3  Sector-Coordinator 
            writeOneRowToExcel(output, EXCEL_OUTPUT_START_ROW, "Sl.No", "From", "To", "Center", "Nature of Activity", persons, sectorCoordinatorTitle);
        }

        private void writeOneRowToExcel(final WritableSheet output, final int row, final String slNo, final String from,
                final String to, final String place, final String activity, final List<String> persons,
                final String placeOwner) throws RowsExceededException, WriteException, ParseException {

            int col = EXCEL_OUTPUT_START_COL;
            {
                if(Character.isDigit(slNo.charAt(0))) {
                    jxl.write.Number number = new jxl.write.Number(col++, row, Integer.parseInt(slNo));
                    output.addCell(number);
                }
                else {
                    Label label = new Label(col++, row, slNo);
                    output.addCell(label);
                }
            }
            {
                if(Character.isDigit(from.charAt(0))) {
                    col = writeDateToExcel(output, col, row, from);
                }
                else {
                    Label label = new Label(col++, row, from);
                    output.addCell(label);
                }
            }
            {
                if(Character.isDigit(from.charAt(0))) {
                    col = writeDateToExcel(output, col, row, to);
                }
                else {
                    Label label = new Label(col++, row, to);
                    output.addCell(label);
                }
            }
            {
                Label label = new Label(col++, row, place);
                output.addCell(label);
            }
            {
                String unmaskedActivity = unmaskFromGrouping(activity);
                Label label = new Label(col++, row, unmaskedActivity);
                output.addCell(label);
            }
            int placeOwnerCol = col + maxNumberOfPersons;
            {
                for(String person : persons) {
                    Label label = new Label(col++, row, person);
                    output.addCell(label);
                }
            }
            {
                Label label = new Label(placeOwnerCol++, row, placeOwner);
                output.addCell(label);
            }
        }

        //If these are done for each of the rows, it results in a warning like below.
        //Warning:  Maximum number of format records exceeded.  Using default format.
        //http://support.teamdev.com/thread/1760
        private WritableCellFormat dateFormat = new WritableCellFormat(new DateFormat("dd-MMM-yyyy"));

        private int writeDateToExcel(final WritableSheet output, int col, final int row, final String date)
                throws ParseException, WriteException, RowsExceededException {
            Date now = scheduleDateFormat.parse(date);
            DateTime dateCell = new DateTime(col++, row, now, dateFormat);
            output.addCell(dateCell);

            return col;
        }

        //For each activity for the given teacher, 
        //update the maps and return schedule for this teacher.
        //scheduleStartDate and scheduleEndDate if given, schedule will be prepared only 
        //for activities that fall completely under these two dates inclusive of both.
        //If anyone of them is null or empty, they are ignored.
        //If scheduleStartDate is after scheduleEndDate, an exception is thrown.
        private void fillUpPlaceActivityMap(
                final Map<String, Map<String, Map<String, Map<String, List<String>>>>> startEndPlaceActivityPersonMap,
                final String teacher, 
                final Sheet inputSheet, 
                final int teacherScheduleRow,
                final String scheduleStartDate,
                final String scheduleEndDate) {

            Map<String, Map<String, List<String>>> placeActivityToDatesMap = 
                new HashMap<String, Map<String,List<String>>>();

            int startColumn = TEACHER_START_COL + 1;
            int endColumn = inputSheet.getColumns();
            String prevPlace = "";
            String prevActivity = "";
            for(int col = startColumn; col < endColumn; col++) {
                Cell placeActivityCell = inputSheet.getCell(col, teacherScheduleRow);
                String placeActivity = placeActivityCell.getContents().trim();

                if("".equals(placeActivity)) {
                    //nothing to do. find the next place activity string.
                    continue;
                }

                // New place activity found. Set the end date for the previous place-activity
                setEndDate(startEndPlaceActivityPersonMap, 
                        placeActivityToDatesMap, prevPlace, prevActivity, teacher, col - 1);

                // placeActivity string can be of three forms:
                // 1. Delhi - Training (or without leading / trailing space for the hypen)
                // 2. Center [/ sector] - activity (the optional sector will be used in the place owner map)
                // 3. BREAK
                // 4. Travel
                String place = "";
                String activity = placeActivity;
                int separatorIndex = placeActivity.indexOf(PLACE_ACTIVITY_SEPARATOR);
                if(separatorIndex > 0) {
                    place = placeActivity.substring(0, separatorIndex).trim();
                    activity = placeActivity.substring(separatorIndex + 1).trim();
                }

                // Find the start date for this new place-activity.
                String startDate = dateHelper.getDate(col);
                if(!"".equals(scheduleStartDate) && scheduleDateComparator.compare(startDate, scheduleStartDate) < 0) {
                    if(isDebug)
                        System.out.println("Skipping schedule as the start date: " + startDate + " occurs before scheduleStartDate: " + scheduleStartDate);
                    continue;
                }
                List<String> startEndDates = new ArrayList<String>();
                startEndDates.add(startDate);
                // If possible, get the end date as well from the merged cells
                int endCol = -1;
                Map<CellInfo, String> valueMap = MERGED_CELLS_MAP.get(new CellInfo(placeActivityCell));
                if(valueMap != null) {
                    for(CellInfo bottomRightInfo : valueMap.keySet()) {
                        endCol = bottomRightInfo.col;
                    }
                }
                else {
                    //If the current cell is not a merged-cell, the only choice it has is it being a single
                    //date activity. In that case, set the end date as the current date itself.
                    endCol = col;
                }
                String endDate = dateHelper.getDate(endCol);
                if(!"".equals(scheduleEndDate) && scheduleDateComparator.compare(endDate, scheduleEndDate) > 0) {
                    if(isDebug)
                        System.out.println("Skipping schedule as the end date: " + endDate + " occurs after scheduleEndDate: " + scheduleEndDate);
                    continue;
                }
                startEndDates.add(endDate);

                Map<String, List<String>> activityMap = new HashMap<String, List<String>>();
                activityMap.put(activity, startEndDates);
                placeActivityToDatesMap.put(place, activityMap);

                // Set the previous placeActivity
                prevPlace = place;
                prevActivity = activity;
            }

            setEndDate(startEndPlaceActivityPersonMap, 
                    placeActivityToDatesMap, prevPlace, prevActivity, teacher, endColumn - 1);

            if(isDebug) System.out.println("Place activity map for teacher: " + teacher + ": " + placeActivityToDatesMap);
        }

        private void setEndDate(
                final Map<String, Map<String, Map<String, Map<String, List<String>>>>> startEndPlaceActivityPersonMap,
                final Map<String, Map<String, List<String>>> placeActivityMap, 
                final String place, 
                final String activity, 
                final String teacher, 
                final int endColumn) {
            if(!"".equals(place) || !"".equals(activity)) {
                List<String> startEndDates = placeActivityMap.get(place).get(activity);
                //do not add end date if it is already added.
                //it will be already added if it was part of the merged cell.
                if(startEndDates.size() == 1) {
                    startEndDates.add(dateHelper.getDate(endColumn));
                }
                placeActivityMap.get(place).put(activity, startEndDates);

                updateStartEndPlaceActivityPersonMap(startEndPlaceActivityPersonMap, 
                        startEndDates.get(0), startEndDates.get(1), place, activity, teacher);
            }
        }

        private void updateStartEndPlaceActivityPersonMap(
                final Map<String, Map<String, Map<String, Map<String, List<String>>>>> startEndPlaceActivityPersonMap,
                final String startDate,
                final String endDate,
                final String place,
                String activity,
                final String teacher) {
            Map<String, Map<String, Map<String, List<String>>>> endPlaceActivityMap = 
                startEndPlaceActivityPersonMap.get(startDate);
            if(endPlaceActivityMap == null) {
                endPlaceActivityMap = new TreeMap<String, Map<String, Map<String, List<String>>>>(scheduleDateComparator);
                startEndPlaceActivityPersonMap.put(startDate, endPlaceActivityMap);
            }

            Map<String, Map<String, List<String>>> placeActivityMap = endPlaceActivityMap.get(endDate);
            if(placeActivityMap == null) {
                placeActivityMap = new LinkedHashMap<String, Map<String,List<String>>>();
                endPlaceActivityMap.put(endDate, placeActivityMap);
            }

            Map<String, List<String>> activityMap = placeActivityMap.get(place);
            if(activityMap == null) {
                activityMap = new LinkedHashMap<String, List<String>>();
                placeActivityMap.put(place, activityMap);
            }

            activity = maskForGrouping(activity, teacher);
            List<String> teachers = activityMap.get(activity);
            if(teachers == null) {
                teachers = new ArrayList<String>();
                activityMap.put(activity, teachers);
            }

            if(!teachers.contains(teacher)) {
                teachers.add(teacher);
                if(teachers.size() > maxNumberOfPersons) 
                    maxNumberOfPersons = teachers.size();
            }
            for(String member : teachers) {
                //Store the max group of teachers for each teacher
                Set<String> groupMembers = teacherToLargestGroupMap.get(member);
                if(groupMembers == null) {
                    groupMembers = new HashSet<String>();
                    teacherToLargestGroupMap.put(member, groupMembers);
                }
                if(groupMembers.size() < teachers.size()) {
                    groupMembers.clear();
                    groupMembers.addAll(teachers);
                }
            }
        }

        // Add teacher to the activity if needed.
        private String maskForGrouping(String activity, String teacher) {
            if(!isWhitelistedForGrouping(activity)) {
                //Do not group any activity across teachers unless its whitelisted explicitly
                //even if they fall on the same date.
                activity = activity + "-" + teacher;
            }
            return activity;
        }

        private boolean isWhitelistedForGrouping(String activity) {
            String[] tokens = activity.split(" ");
            for(String token : tokens) {
                if(ACTIVITIES_MARKED_FOR_GROUPING.contains(token.trim().toLowerCase())) {
                    return true;
                }
            }
            return false;
        }

        // Remove teacher if present.
        private String unmaskFromGrouping(String activity) {
            int separatorIndex = activity.lastIndexOf("-");
            if(separatorIndex >= 0) {
                activity = activity.substring(0, separatorIndex);
            }
            return activity;
        }
    }
}
