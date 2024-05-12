import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class FRS {

    // TODO : Global variables to read Excel columns

    // For reading Application Form; Input the column in that sheet
    // private static final int STUDENT_NAME_COLUMN = 2;
    private static final int STUDENT_ID_COLUMN = 2; // 3 - 1 = 2
    private static final int RETAKE_COURSE_COLUMN = 4; // 5 - 1 = 4

    // For reading Class Details; Input the column in that sheet
    private static final int COURSE_ID_COLUMN = 0; // 1 - 1 = 0
    private static final int COURSE_CLASS_COLUMN = 6; // 7 - 1 = 6
    private static final int NUMBER_OF_STUDENTS_COLUMN = 7; // 8 - 1 = 7
    private static final int DEPARTMENTS_COLUMN = 8; // 9 - 1 = 8

    // To write the class number
    private static final int WRITE_CLASS_NUMBER_COLUMN = 5; // 6 - 1 = 5
    private static final int NOTE_COLUMN = 7; // 8 - 1 = 7
    
    private static final int COURSE_LIMIT = 2;
    
    private static final int[] ALL_DEPARTMENT_CODES_IN_ARRAY = 
            {
                5001, 5002, 5003, 5004, 5005, 5006, 5007, 5008, 5009, 5010,
                5011, 5012, 5013, 5014, 5015, 5016, 5017, 5018, 5019, 5020,
                5021, 5022, 5023, 5024, 5025, 5026, 5027, 5028, 5029, 5030,
                5031, 5032, 5033, 5034, 5035, 5036, 5037, 5038, 5039, 5040,
                5041, 5042, 5043, 5044, 5045, 5046, 5047, 5048, 5049, 5050,
                5051, 5052, 5053, 5054, 5055
            };
    
    private static final int PAUSE_VIEW = 1000;

    // private static Map<String, Map<String, Map<Integer, Integer>>> AVAILABLE_CLASSES_FOR_DEPARTMENTS_BY_COURSES = new HashMap<>();
    // private static List<Student> STUDENT_LIST = new ArrayList<>();

    public static void main(String[] args) {
        try {
            System.out.println("> Memulai proses memasukkan data mahasiswa ke dalam kelas-kelas.");
                    
            // TODO: Working the whole process here
            // TODO: Set path for files
            String INPUT_PATHFILE = "C:\\Users\\natha\\Downloads\\Data FRS 2.xlsx";
            String CLASS_DETAILS_PATHFILE = "C:\\Users\\natha\\Downloads\\Class Details Reguler.xlsx";

            // TODO: Setup the working procedures
            // Step 1: Read the Application From
            List<Student> STUDENT_LIST = READ_APPLICATION_FORM(INPUT_PATHFILE, 
                    STUDENT_ID_COLUMN, RETAKE_COURSE_COLUMN, COURSE_ID_COLUMN, 
                    COURSE_CLASS_COLUMN, DEPARTMENTS_COLUMN);
            
//            Thread.sleep(PAUSE_VIEW);

            // Step 2: Read the Class' Details
            Map<String, Map<String, Map<Integer, Integer>>> AVAILABLE_CLASSES_FOR_DEPARTMENTS_BY_COURSES = READ_CLASS_DETAILS(CLASS_DETAILS_PATHFILE, STUDENT_LIST);
            
//            Thread.sleep(PAUSE_VIEW);

            // Step 3: Assign classes for all students
            PROCESS_RESULT MAP_AND_LIST = ASSIGN_CLASS(STUDENT_LIST, AVAILABLE_CLASSES_FOR_DEPARTMENTS_BY_COURSES);
            
//            Thread.sleep(PAUSE_VIEW);

            // Step 3.1: Take the map and student list
            AVAILABLE_CLASSES_FOR_DEPARTMENTS_BY_COURSES = MAP_AND_LIST.getUpdatedAvailableClassesForDeparmentsByCourses();
            STUDENT_LIST = MAP_AND_LIST.getUpdatedStudentList();

            // Step 4: Print to Excel
            WRITE_TO_EXCEL(INPUT_PATHFILE, STUDENT_LIST);
            
//            Thread.sleep(PAUSE_VIEW);
            
        } catch (IOException e) {
            // TODO: Exception for the main program
            System.err.println("> Kesalahan saat menyiapkan file input atau output: " + e.getMessage());
            e.printStackTrace();
        } catch (Exception e) {
            System.err.println("> Terjadi kesalahan: " + e.getMessage());
            e.printStackTrace();
        } finally {
            System.out.println("> Program telah selesai dieksekusi.");
        }
    }

    public static List<Student> READ_APPLICATION_FORM(String PATHFILE, int STUDENT_ID_COLUMN,
            int RETAKE_COURSE_COLUMN, int COURSE_ID_COLUMN, int COURSE_CLASS_COLUMN, int DEPARTMENTS_COLUMN) throws IOException {
        List<Student> STUDENT_LIST = new ArrayList<>();

        // TODO: Set up path to Application Form's Excel and read the file
        try (FileInputStream FILEINPUTSTREAM = new FileInputStream(new File(PATHFILE)); // Read the Excel File
                Workbook WORKBOOK = new XSSFWorkbook(FILEINPUTSTREAM)) {

            // Read the first sheet which contains all the data needed
            Sheet SHEET = WORKBOOK.getSheetAt(0);
            
            int ROW_INDEX = 1;
            boolean IS_DATA_AVAILABLE = true;
            while (IS_DATA_AVAILABLE) {
                // Get the row within the sheet
                Row ROW = SHEET.getRow(ROW_INDEX);
                DataFormatter FORMATTER = new DataFormatter();

                if (ROW == null) {
                    IS_DATA_AVAILABLE = false;
                } else {
                    // Take student's data
                    // If Student's name is necessary, uncomment below
//                    String STUDENT_NAME = ROW.getCell(STUDENT_NAME_COLUMN).getStringCellValue(); 
                    Cell THIS_CELL = ROW.getCell(STUDENT_ID_COLUMN);
                    String STUDENT_ID = FORMATTER.formatCellValue(THIS_CELL);
                    if (STUDENT_ID == "") {
                        break;
                    }

                    String STUDENT_DEPARTMENT = STUDENT_ID.substring(0, 4);
                    String RETAKE_COURSE = 
                        ROW.getCell(RETAKE_COURSE_COLUMN).getStringCellValue().substring(0, 2) + 
                        "23" + 
                        ROW.getCell(RETAKE_COURSE_COLUMN).getStringCellValue().substring(2);

                    // Put into Student Object
                    Student STUDENT = new Student(STUDENT_ID, STUDENT_DEPARTMENT, RETAKE_COURSE);
                    STUDENT_LIST.add(STUDENT);
                    ROW_INDEX++;
                }
            }
            
            System.out.println("> Berhasil membaca data mahasiswa.");

        } catch (FileNotFoundException e) {
            // TODO: handle exception
            System.err.println("> File tidak ditemukan: " + e.getMessage());
            e.printStackTrace();
        } catch (IOException e) {
            System.err.println("> Kesalahan saat membaca file: " + e.getMessage());
            e.printStackTrace();
        } catch (Exception e) {
            System.err.println("> Terjadi kesalahan: " + e.getMessage());
            e.printStackTrace();
        }

        return STUDENT_LIST;
    }

    public static Map<String, Map<String, Map<Integer, Integer>>> READ_CLASS_DETAILS(String PATHFILE, List<Student> STUDENT_LIST) {
        // TODO: Generate a map as a return value
        Map<String, Map<String, Map<Integer, Integer>>> AVAILABLE_CLASSES_FOR_DEPARTMENTS_BY_COURSES = new HashMap<>();

        // TODO: Read Class Details'
        // Load the Excel file within try-catch statement
        try (FileInputStream FILEINPUTSTREAM = new FileInputStream(new File(PATHFILE));
            Workbook WORKBOOK = new XSSFWorkbook(FILEINPUTSTREAM)) {

            for (Student STUDENT : STUDENT_LIST) {
                // TODO: Search for the letter in the course code and find the corresponding sheet
                // Get the sheet
                String COURSE_SHEET = STUDENT.getRetakeCourse();
                Sheet SHEET = WORKBOOK.getSheet(COURSE_SHEET);
                
                // Remove iteratived checking; ignore similar elements (ctx: COURSE_CODE checking)
                if (COURSE_SHEET.equals(AVAILABLE_CLASSES_FOR_DEPARTMENTS_BY_COURSES.containsKey(COURSE_SHEET))) {
                    continue;
                }
                
                if (SHEET != null) {
                    int ROW_INDEX = 4;
                    boolean IS_DATA_AVAILABLE = true;
                    
                    while (IS_DATA_AVAILABLE) {
                        Row ROW = SHEET.getRow(ROW_INDEX);
                        
                        // Safety measures to avoid null cells
                        boolean HAS_DATA = false;
                        
                        if (ROW != null && ROW.getFirstCellNum() > -1) {
                            for (int CELL_INDEX = ROW.getFirstCellNum(); CELL_INDEX < 11; CELL_INDEX++) {
                                if (CELL_INDEX < 0) {
                                    break;
                                }

                                Cell THE_CELL = ROW.getCell(CELL_INDEX, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

                                if (THE_CELL.getCellTypeEnum() == CellType.STRING || THE_CELL.getCellTypeEnum() == CellType.NUMERIC) {
                                    HAS_DATA = true;
                                    break;
                                } else {
                                    break;
                                }
                            }

                            if (HAS_DATA) {
                                Cell CLASS_CELL = ROW.getCell(COURSE_CLASS_COLUMN);
                                Cell ENROLLED_CELL = ROW.getCell(NUMBER_OF_STUDENTS_COLUMN);
                                Cell DEPARTMENTS_CELL = ROW.getCell(DEPARTMENTS_COLUMN);

                                String COURSE_CODE = COURSE_SHEET;
                                int COURSE_CLASS = (int) CLASS_CELL.getNumericCellValue();
                                int NUMBER_OF_STUDENTS = (int) ENROLLED_CELL.getNumericCellValue();

                                String DEPARTMENT_CODES = DEPARTMENTS_CELL.getStringCellValue();
                                String[] DEPARTMENT_CODES_IN_ARRAY = DEPARTMENT_CODES.split(",\\s*");

                                for (String DEPARTMENT : DEPARTMENT_CODES_IN_ARRAY) {
                                    String THE_DEPARTMENT = DEPARTMENT;
                                    if (DEPARTMENT.equals("Semua Departemen") || DEPARTMENT.equals("semua Dep (mhs ulang)")) {
                                        DEPARTMENT_CODES_IN_ARRAY = DEPARTMENT_CODES_IN_STRING(ALL_DEPARTMENT_CODES_IN_ARRAY);
                                        for (String CODES : DEPARTMENT_CODES_IN_ARRAY) {
                                            AVAILABLE_CLASSES_FOR_DEPARTMENTS_BY_COURSES
                                                .computeIfAbsent(COURSE_CODE, k -> new HashMap<>())
                                                .computeIfAbsent(CODES, k -> new HashMap<>())
                                                .put(COURSE_CLASS, NUMBER_OF_STUDENTS);
                                        }
                                        break;
                                    } else {
                                        AVAILABLE_CLASSES_FOR_DEPARTMENTS_BY_COURSES
                                            .computeIfAbsent(COURSE_CODE, k -> new HashMap<>())
                                            .computeIfAbsent(THE_DEPARTMENT, k -> new HashMap<>())
                                            .put(COURSE_CLASS, NUMBER_OF_STUDENTS);
                                    }
                                }
                                ROW_INDEX++;
                            } else {
                                break;
                            }
                        } else {
                            IS_DATA_AVAILABLE = false;
                            break;
                        }
                    }
                }
            }
            
            System.out.println("> Berhasil membaca detail kelas semua mata kuliah.");
        } catch (NullPointerException e) {
            System.err.println("> Data null muncul saat membaca detail kelas: " + e.getMessage());
            e.printStackTrace();
        } catch (FileNotFoundException e) {
            // TODO: handle exception
            System.err.println("> File tidak ditemukan saat membaca detail kelas: " + e.getMessage());
            e.printStackTrace();
        } catch (IOException e) {
            System.err.println("> Kesalahan membaca file saat membaca detail kelas: " + e.getMessage());
            e.printStackTrace();
        } catch (Exception e) {
            System.err.println("> Terjadi kesalahan saat membaca detail kelas: " + e.getMessage());
            e.printStackTrace();
        }

        return AVAILABLE_CLASSES_FOR_DEPARTMENTS_BY_COURSES;
    }

    public static PROCESS_RESULT ASSIGN_CLASS(List<Student> STUDENT_LIST, 
        Map<String, Map<String, Map<Integer, Integer>>> AVAILABLE_CLASSES_FOR_DEPARTMENTS_BY_COURSES) {

        try {
            for (Student STUDENT : STUDENT_LIST) {
                // TODO: Break the main map (AVAILABLE_CLASSES_FOR_DEPARTMENTS_BY_COURSES)
                Map<String, Map<Integer, Integer>> AVAILABLE_CLASSES_FOR_DEPARTMENTS = 
                        AVAILABLE_CLASSES_FOR_DEPARTMENTS_BY_COURSES.getOrDefault(STUDENT.getRetakeCourse(), new HashMap<>());
                Map<Integer, Integer> AVAILABLE_CLASSES = AVAILABLE_CLASSES_FOR_DEPARTMENTS.getOrDefault(STUDENT.getDepartment(), new HashMap<>());

                String DEPARTMENT = STUDENT.getDepartment();
                String RETAKE_COURSE = STUDENT.getRetakeCourse();
                int ASSIGN_TO_THIS_CLASS = GET_THE_LEAST_NUMBER_OF_STUDENT_AMONG_CLASSES(AVAILABLE_CLASSES);
                int NUMBER_OF_STUDENTS = AVAILABLE_CLASSES.getOrDefault(ASSIGN_TO_THIS_CLASS, 0);
                STUDENT.setCourseClass(ASSIGN_TO_THIS_CLASS);

                AVAILABLE_CLASSES.put(ASSIGN_TO_THIS_CLASS, NUMBER_OF_STUDENTS + 1);
                
                UPDATE_NUMBER_OF_STUDENTS(AVAILABLE_CLASSES_FOR_DEPARTMENTS);
                
                AVAILABLE_CLASSES_FOR_DEPARTMENTS.put(DEPARTMENT, AVAILABLE_CLASSES);
                AVAILABLE_CLASSES_FOR_DEPARTMENTS_BY_COURSES.put(RETAKE_COURSE, AVAILABLE_CLASSES_FOR_DEPARTMENTS);
            }
        } catch (Exception e) {
            System.err.println("> Terjadi kesalahan saat ingin memasukkan mahasiswa ke dalam kelas: " + e.getMessage());
            e.printStackTrace();
        }

        return new PROCESS_RESULT(AVAILABLE_CLASSES_FOR_DEPARTMENTS_BY_COURSES, STUDENT_LIST);
    }
    
    public static void UPDATE_NUMBER_OF_STUDENTS(Map<String, Map<Integer, Integer>> AVAILABLE_CLASSES_FOR_DEPARTMENTS) {
        for (Map.Entry<String, Map<Integer, Integer>> ENTRY : AVAILABLE_CLASSES_FOR_DEPARTMENTS.entrySet()) {
            String THE_DEPARTMENT = ENTRY.getKey();
            Map<Integer, Integer> CLASS_MAP = ENTRY.getValue();
            
            for (Map.Entry<Integer, Integer> CLASS_ENTRY : CLASS_MAP.entrySet()) {
                int CLASS_NUMBER = CLASS_ENTRY.getKey();
                int NUMBER_OF_STUDENTS = CLASS_ENTRY.getValue();
                
                for (Map.Entry<String, Map<Integer, Integer>> COMPARING_ENTRY : AVAILABLE_CLASSES_FOR_DEPARTMENTS.entrySet()) {
                    if (!COMPARING_ENTRY.getKey().equals(THE_DEPARTMENT)) {
                        Map<Integer, Integer> COMPARING_CLASS_MAP = COMPARING_ENTRY.getValue();
                        
                        if (COMPARING_CLASS_MAP.containsKey(CLASS_NUMBER)) {
                            int COMPARING_NUMBER_OF_STUDENTS = COMPARING_CLASS_MAP.get(CLASS_NUMBER);
                            
                            if (COMPARING_NUMBER_OF_STUDENTS > NUMBER_OF_STUDENTS) {
                                CLASS_MAP.put(CLASS_NUMBER, COMPARING_NUMBER_OF_STUDENTS);
                                NUMBER_OF_STUDENTS = COMPARING_NUMBER_OF_STUDENTS;
                            }
                        }
                    }
            }
            }
        }
    }

    public static int GET_THE_LEAST_NUMBER_OF_STUDENT_AMONG_CLASSES(Map<Integer, Integer> AVAILABLE_CLASSES) {
        List<Map.Entry<Integer, Integer>> SORTING_ENTRIES_LIST = new ArrayList<>(AVAILABLE_CLASSES.entrySet());
        
        if (SORTING_ENTRIES_LIST.isEmpty()) {
            return 0;
        }
        
        SORTING_ENTRIES_LIST.sort(Map.Entry.comparingByValue());
        Map.Entry<Integer, Integer> LEAST_STUDENT = SORTING_ENTRIES_LIST.get(0);
        
        return LEAST_STUDENT.getKey();
    }

    public static void WRITE_TO_EXCEL(String PATHFILE, List<Student> STUDENT_LIST) throws FileNotFoundException, IOException, InvalidFormatException {
        try (FileInputStream OUTPUT_FILE = new FileInputStream(PATHFILE); 
                Workbook OUTPUT_WORKBOOK = WorkbookFactory.create(OUTPUT_FILE)) {
            Sheet OUTPUT_SHEET = OUTPUT_WORKBOOK.getSheetAt(0);
            
            Map<String, Integer> ID_COUNT = new HashMap<>();

            int OUTPUT_ROW_NUMBER = 1;
            System.out.println();
            for (Student STUDENTS : STUDENT_LIST) {
                String STUDENT_ID = STUDENTS.getStudentID();
                int COUNT = ID_COUNT.getOrDefault(STUDENT_ID, 0);
                ID_COUNT.put(STUDENT_ID, COUNT + 1);
                
                Row OUTPUT_ROW = OUTPUT_SHEET.getRow(OUTPUT_ROW_NUMBER);
                
                if (OUTPUT_ROW == null) {
                    OUTPUT_ROW = OUTPUT_SHEET.createRow(OUTPUT_ROW_NUMBER);
                }

                OUTPUT_ROW.createCell(WRITE_CLASS_NUMBER_COLUMN).setCellValue(STUDENTS.getAssignedClass());
                    if (STUDENTS.getAssignedClass() == 0) {
                        System.out.println("> Mahasiswa dengan NRP " + STUDENTS.getStudentID() + " pada mata kuliah " + STUDENTS.getRetakeCourse() + " tidak memiliki kelas, periksa file output.");
                    }
                    
                if (COUNT > (COURSE_LIMIT - 1)) {
                    OUTPUT_ROW.createCell(NOTE_COLUMN).setCellValue("Hapus");
                }
                
                OUTPUT_ROW_NUMBER++;
            }
            
            System.out.println();
            for (Map.Entry<String, Integer> ENTRY_ID_COUNT : ID_COUNT.entrySet()) {
                String ENTRY_ID = ENTRY_ID_COUNT.getKey();
                int COUNT = ENTRY_ID_COUNT.getValue();
                
                if (COUNT > COURSE_LIMIT) {
                    System.out.println("> Mahasiswa dengan NRP " + ENTRY_ID + " mendaftar sebanyak " + COUNT + " mata kuliah, periksa file output.");
                }
            }
            
            
            try (FileOutputStream OUT_FILE = new FileOutputStream(PATHFILE)) {
                OUTPUT_WORKBOOK.write(OUT_FILE);
            }
//            Thread.sleep(PAUSE_VIEW);
            System.out.println("""
                               
                               > Mahasiswa telah dimasukkan ke kelas-kelas.""");

        } catch (FileNotFoundException e) {
            // TODO: handle exception
            System.err.println("> File tidak ditemukan saat ingin mencetak hasil program: " + e.getMessage());
            e.printStackTrace();
        } catch (IOException e) {
            System.err.println("> Kesalahan membaca file saat ingin mencetak hasil program: " + e.getMessage());
            e.printStackTrace();
        } catch (Exception e) {
            System.err.println("> Kesalahan terjadi saat ingin mencetak hasil program: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    public static String[] DEPARTMENT_CODES_IN_STRING(int[] ALL_DEPARTMENT_CODES_IN_ARRAY) {
        String[] ALL_DEPARTMENT_CODES_IN_STRING_ARRAY = 
                Arrays.stream(ALL_DEPARTMENT_CODES_IN_ARRAY)
                .mapToObj(String::valueOf)
                .toArray(String[]::new);
        return ALL_DEPARTMENT_CODES_IN_STRING_ARRAY;
    }

    // Class student to access objects
    private static class Student {
        private String STUDENT_NAME;
        private String STUDENT_ID;
        private String STUDENT_DEPARTMENT;

        // We make this individually to fit given Excel file data
        private String RETAKE_COURSE;
        private int RETAKE_COURSE_CLASS;

        public Student(String STUDENT_ID, String STUDENT_DEPARTMENT, String RETAKE_COURSE) {
            this.STUDENT_ID = STUDENT_ID;
            this.STUDENT_DEPARTMENT = STUDENT_DEPARTMENT;
            this.RETAKE_COURSE = RETAKE_COURSE;
        }
        public Student(String STUDENT_NAME, String STUDENT_ID, String STUDENT_DEPARTMENT, String RETAKE_COURSE) {
            this.STUDENT_NAME = STUDENT_NAME;
            this.STUDENT_ID = STUDENT_ID;
            this.STUDENT_DEPARTMENT = STUDENT_DEPARTMENT;
            this.RETAKE_COURSE = RETAKE_COURSE;
        }

        public String getName() {
            return STUDENT_NAME;
        }

        public String getStudentID() {
            return STUDENT_ID;
        }

        public String getRetakeCourse() {
            return RETAKE_COURSE;
        }

        public String getDepartment() {
            return STUDENT_DEPARTMENT;
        }

        public int getAssignedClass() {
            return RETAKE_COURSE_CLASS;
        }

        public void setCourseClass(int RETAKE_COURSE_CLASS) {
            this.RETAKE_COURSE_CLASS = RETAKE_COURSE_CLASS;
        }
    }

    public static class PROCESS_RESULT {
        private final Map<String, Map<String, Map<Integer, Integer>>> UPDATED_AVAILABLE_CLASSES_FOR_DEPARTMENTS_BY_COURSES;
        private final List<Student> UPDATED_STUDENT_LIST;

        public PROCESS_RESULT(
            Map<String, Map<String, Map<Integer, Integer>>> UPDATED_AVAILABLE_CLASSES_FOR_DEPARTMENTS_BY_COURSES,
            List<Student> UPDATED_STUDENT_LIST
        ) {
            this.UPDATED_AVAILABLE_CLASSES_FOR_DEPARTMENTS_BY_COURSES = UPDATED_AVAILABLE_CLASSES_FOR_DEPARTMENTS_BY_COURSES;
            this.UPDATED_STUDENT_LIST = UPDATED_STUDENT_LIST;
        }

        public Map<String, Map<String, Map<Integer, Integer>>> getUpdatedAvailableClassesForDeparmentsByCourses() {
            return UPDATED_AVAILABLE_CLASSES_FOR_DEPARTMENTS_BY_COURSES;
        }

        public List<Student> getUpdatedStudentList() {
            return UPDATED_STUDENT_LIST;
        }
    }
}