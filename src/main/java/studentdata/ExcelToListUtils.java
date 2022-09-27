package studentdata;

import java.io.IOException;
import java.util.*;

import org.apache.poi.openxml4j.exceptions.InvalidOperationException;
import java.io.File;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

class Student {

	String name;
	int admissionNo;
	float percentage;
	int physicsMark;
	int chemistryMark;
	int mathsMark;
	static Logger logger = LogManager.getLogger(Student.class);

	Student(String name, int admissionNo, float percentage, int physicsMark, int chemistryMark, int mathsMark) {

		this.name = name;
		this.admissionNo = admissionNo;
		this.percentage = percentage;
		this.physicsMark = physicsMark;
		this.chemistryMark = chemistryMark;
		this.mathsMark = mathsMark;

	}

	public void display(String physicsGrade, String chemistryGrade, String mathsGrade, Object physicsGradePointt,
			Object chemistryGradePoint, Object mathsGradePoint) {

		String grade = "Grade";
		String gradePoint = "Grade Point";
		String format = "\t{}: {}";
		logger.info("");
		logger.info("Name: {}", this.name);
		logger.info("Admission No: {}", this.admissionNo);
		logger.info("Percentage: {}", this.percentage);
		logger.info("Physics: ");
		logger.info("\t Mark: {}", this.physicsMark);
		logger.info(format, grade, physicsGrade);
		logger.info(format, gradePoint, physicsGradePointt);
		logger.info("Chemistry: ");
		logger.info("\tMark: {}", this.chemistryMark);
		logger.info(format, grade, chemistryGrade);
		logger.info(format, gradePoint, chemistryGradePoint);
		logger.info("Maths: ");
		logger.info("\tMark: {}", this.mathsMark);
		logger.info(format, grade, mathsGrade);
		logger.info(format, gradePoint, mathsGradePoint);
		logger.info("");

	}

	public static void excelToListConverter() throws IOException {

		try {

			Scanner scanner = new Scanner(System.in);
			File file = new File("C:\\Users\\ecs\\Desktop\\ECS training\\resources\\" + "Book1.xlsx");
			String path = file.getAbsolutePath();

			ExcelUtils excel = new ExcelUtils();
			excel.strcellData(0, 1, path);

			int rowCount = excel.rowCount(path) - 1;

			// creation of arraylist for student details

			List<Student> list = new ArrayList<Student>();

			String[] name = new String[rowCount];
			int[] admissionNo = new int[rowCount];
			float[] percentage = new float[rowCount];

			int[] physicsMark = new int[rowCount];
			String[] physicsGrade = new String[rowCount];
			Object[] physicsGradePoint = new Object[rowCount];

			int[] chemistryMark = new int[rowCount];
			String[] chemistryGrade = new String[rowCount];
			Object[] chemistryGradePoint = new Object[rowCount];

			int[] mathsMark = new int[rowCount];
			String[] mathsGrade = new String[rowCount];
			Object[] mathsGradePoint = new Object[rowCount];

			float[] total = new float[rowCount];

			// adding datas to the arraylist created

			for (int i = 0; i < rowCount; i++) {

				name[i] = excel.strcellData(i + 1, 1, path);
				admissionNo[i] = excel.numcellData(i + 1, 0, path);
				physicsMark[i] = excel.numcellData(i + 1, 2, path);
				chemistryMark[i] = excel.numcellData(i + 1, 3, path);
				mathsMark[i] = excel.numcellData(i + 1, 4, path);
				total[i] = physicsMark[i] + chemistryMark[i] + (float) mathsMark[i];
				percentage[i] = (total[i] * 100) / 300;

				physicsGrade[i] = excel.gradeCalc(physicsMark[i]);
				gradeAssigner(physicsMark, physicsGradePoint, i);

				chemistryGrade[i] = excel.gradeCalc(chemistryMark[i]);
				gradeAssigner(chemistryMark, chemistryGradePoint, i);

				mathsGrade[i] = excel.gradeCalc(mathsMark[i]);
				gradeAssigner(mathsMark, mathsGradePoint, i);

			}

			for (int j = 0; j < rowCount; j++) {

				Student s = new Student(name[j], admissionNo[j], percentage[j], physicsMark[j], chemistryMark[j],
						mathsMark[j]);
				list.add(s);

			}

			logger.info("Type \"name\" to search by name or type \"admission\" to search by admission number : ");
			String chooser = scanner.nextLine();

			if (chooser.equals("name")) {

				logger.info("Type the name you want to search :");
				String searchName = scanner.nextLine();

				for (int k = 0; k < list.size(); k++) {

					if (searchName.equals(name[k])) {

						list.get(k).display(physicsGrade[k], chemistryGrade[k], mathsGrade[k], physicsGradePoint[k],
								chemistryGradePoint[k], mathsGradePoint[k]);
					}

				}

			}

			if (chooser.equals("admission")) {

				logger.info("Type the admission number you want to search :");
				String addmissionNumber = scanner.nextLine();
				int admissionNum = Integer.parseInt(addmissionNumber);

				for (int m = 0; m < list.size(); m++) {

					if (admissionNo[m] == admissionNum) {

						list.get(m).display(physicsGrade[m], chemistryGrade[m], mathsGrade[m], physicsGradePoint[m],
								chemistryGradePoint[m], mathsGradePoint[m]);

					}
				}
			}
			scanner.close();
		} catch (InvalidOperationException e) {

			logger.info("The path you entered doesn't exist or the file may not be an excel");
		}

	}

	public static void gradeAssigner(int[] mark, Object[] gradePoint, int i) {

		ExcelUtils excel = new ExcelUtils();
		if (mark[i] < 32) {

			gradePoint[i] = "C";

		} else {

			gradePoint[i] = excel.gradePointCalc(mark[i]);

		}

	}

}
