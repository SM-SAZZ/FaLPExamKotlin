import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File

class Student(val group: String, // ������
              val name: String, // ���
              val ex: Int, // ��� �����
              val lecAttendance: List<Boolean>, // ������������ ������
              val labAttendance: List<Boolean>, // ������������ ���
              val labWorks: List<Int>, // ������ �� ��
              val exams: List<Int>, // ������ �� ��
              val tests: List<Int>, // ������ �� ��
              val attest: String) { // ����������
// ����� ��������
    fun write() {
        println("�������: $name\n" +
                "��� �����: $ex\n" +
                "������������: ${attendanceGrade()}\n" +
                "������ ��:\n" +
                "\t������������ ������: $labWorks\n" +
                "\t��������������� ������: $exams\n" +
                "\t����������� ������: $tests\n" +
                "����� ������: ${getTotalScore() + ex}\n")
    }
// ������� ������ ������������
    fun attendanceGrade(): Int {
        val lecTrueCount = lecAttendance.count { it }
        val labTrueCount = labAttendance.count { it }
        val attendScore = lecTrueCount.toDouble() / lecAttendance.size * 5 + labTrueCount.toDouble() / labAttendance.size * 5
        return when {
            attendScore < 4 -> 0
            attendScore < 6 -> 4
            attendScore < 7.5 -> 6
            attendScore < 9 -> 8
            else -> 10
        }
    }
// ������� ���� ������ ��� ������� (��� �� � ��� ������)
    fun getTotalScore(): Int {
        return attendanceGrade() + labWorks.sum() + tests.sum() * 2
    }
// ��������� ������ �� ��������
    fun evalPreGrade(): String {
        return when {
            getTotalScore() < 30 -> "��������"
            (getTotalScore() + ex) < 60 -> "������"
            (getTotalScore() + ex) < 75 -> "�������"
            (getTotalScore() + ex) < 90 -> "������"
            else -> "�������"
        }
    }
}

data class GroupStats(
    val avgAttendance: Double,
    val avgLabWorks: List<Double>,
    val avgExams: List<Double>,
    val avgTests: List<Double>,
    val commonAttest: String,
    val avgCommonLabWorks: Double,
    val avgCommonExams: Double,
    val avgCommonTests: Double,
    val avgCommonAccess: Double,
    val avgCommonTotal: Double,
    val avgPercentLabWorks: Double,
    val avgPercentExams: Double,
    val avgPercentTests: Double,
    val avgPercentAccess: Double,
    val avgPercentTotal: Double
)
// ������ ������� �������� �� Int
fun averageIntList(lists: List<List<Int>>): List<Double> {
    if (lists.isEmpty()) return emptyList()

    val size = lists.first().size
    return List(size) { index ->
        lists.map { it[index] }.average()
    }
}

// ����� �� �� � ��� ������ � ����������� �� ����, ���� ������
fun mostCommonAttest(attests: List<String>): String {
    return attests.groupBy { it }.maxByOrNull { it.value.size }?.key ?: ""
}

// ��������� ������ �� ������. ���� ��� �����, �� ���������� ���, ���� ������, �� ����������� � �����
fun getCellGrade(cell: Cell): Int{
    return when (cell.cellType) {
        CellType.STRING -> {
            val stringValue = cell.stringCellValue
            when (stringValue) {
                "�" -> 2
                "", "-" -> 0
                else -> stringValue.toIntOrNull()
                    ?: 0 // ��������� ������������� ������ � �����, ���� ��� �� �������, �� 0
            }
        }
        else -> cell.numericCellValue.toInt()
    }
}

// ������ �����
fun fileRead(): List<Student> {
    val students = mutableListOf<Student>()

    val file = File("students.xlsx")
    val workbook = WorkbookFactory.create(file.inputStream())
    // ������ �� ���� ������
    for (sheetNum in 0 until workbook.numberOfSheets) {
        val sheet = workbook.getSheetAt(sheetNum)
        var rowC = 0
        // ����� ������ �����
        if (("36" in sheet.sheetName) or ("39" in sheet.sheetName))
            for (row in sheet) {
                // ������� ���� �����
                if (rowC > 1)
                    if (row.getCell(0) != null) {
                        val name = row.getCell(1).stringCellValue // ��������� ���
                        val ex = row.getCell(2).numericCellValue.toInt() // ��������� ��� ������

                        val lecAttendanceList = mutableListOf<Boolean>()
                        val labAttendanceList = mutableListOf<Boolean>()
                        var cellIndex = 4
                        // ������ ������ � ������������ �������� ��� ������ � ��� ���
                        var leclab = true
                        while (sheet.getRow(0).getCell(cellIndex).cellType != CellType.STRING) {
                            // ������������ ����� ������ � ���
                            if (leclab and (sheet.getRow(0).getCell(cellIndex).numericCellValue == 1.0))
                                leclab = false
                            else if (!leclab and (sheet.getRow(0).getCell(cellIndex).numericCellValue == 1.0))
                                leclab = true
                            if (leclab)
                                lecAttendanceList.add(row.getCell(cellIndex).stringCellValue == "+")
                            else
                                labAttendanceList.add(row.getCell(cellIndex).stringCellValue == "+")
                            cellIndex += 1
                        }

                        // ������ ������ �� ������������ ������
                        val labWorks = mutableListOf<Int>()
                        for (i in cellIndex..(cellIndex + 6)) { // ������� � ��1 �� ��7
                            val value = getCellGrade(row.getCell(i))
                            labWorks.add(value)
                        }
                        cellIndex += 7

                        // ������ ������ �� ��������������� ������
                        val exams = mutableListOf<Int>()
                        for (i in cellIndex..(cellIndex + 4)) { // ������� � ��1 �� ��5
                            val value = getCellGrade(row.getCell(i))
                            exams.add(value)
                        }
                        cellIndex += 5

                        // ������ ������ �� ����������� ������
                        val tests = mutableListOf<Int>()
                        for (i in cellIndex..(cellIndex + 2)) { // ������� � ��1 �� ��3
                            val value = getCellGrade(row.getCell(i))
                            tests.add(value)
                        }
                        cellIndex += 32

                        // �������� ������� Student � ���������� ��� � ������
                        val student = Student(sheet.sheetName, name, ex, lecAttendanceList, labAttendanceList, labWorks, exams, tests, row.getCell(cellIndex).stringCellValue)
                        students.add(student)
                    } else break
                rowC += 1
            }
    }

    workbook.close()

    return students
}

// ����� � ������� ����� �������� ��, �� � ��
fun printFormatted(value: Double): String {
    return when {
        value < 0.5 -> "-".padEnd(4)
        value < 1 -> "+".padEnd(4)
        value < 1.5 -> "++".padEnd(4)
        value < 2.5 -> "�".padEnd(4)
        else -> "${String.format("%.1f", value).padEnd(3)} "
    }
}

fun main() {
    /**
     * ���� ������� � ������������ �������� ����� ������ � ��������. ����������
    ��������� ��� ������ � ������������ � ��� ������ �� ��, �� � ��. ������� ��������� �������
    ��� �������� ����������.
    1 ��������� ��������� ����� � ������. ����� ������������ ������, �������
    ��������� ���������� ���������� ������� �� ����� ������ ����� ������������ ������.
    ����� ��, ������� ��������� ���������� ���������� ������� ���� �� �� ������ ����� ��
    �������� ������.
     */
    val students = fileRead()

    val stoopid = students.filter { it.getTotalScore() < 30 }
    var max = 0
    var min = 60
    var maxL = 1
    var minK = 1
    for (i in 0..6) {
        var count = 0
        stoopid.forEach{
            count += if (it.labWorks.get(i) > 0) 1 else 0
        }
        if (count > max) {
            max = count
            maxL = i + 1
        }
    }
    for (i in 0..2) {
        var count = 0
        stoopid.forEach{
            count += if (it.tests.get(i) > 0) 1 else 0
        }
        if (count < min) {
            min = count
            minK = i + 1
        }
    }
    println("������������ ������ �$maxL ���� ����� ���������� ����������� ��������� ��� �������")
    println("����������� ������ �$minK ���� ����� ���������� ����������� ��������� ��� �������\n")

    /**
     * 2 ���������� �������������� � ��������� ����� �� ������������, ��, �� � �� ��� ��,
    ��� � ���������. ����� ����� ������ � ������� ���������� ������ �� ������������, ��, ��,
    ��, ����� ��� ������� ��������.
     */
    students.forEach {
        println("�������: ${it.name}\n" +
                "��� �����: ${it.ex}\n" +
                "������������: ${it.attendanceGrade()}\n" +
                "������ ��:\n" +
                "\t������������ ������: ${it.labWorks}\t�����: ${it.labWorks.sum()}\n" +
                "\t��������������� ������: ${it.exams}\t�����: ${it.exams.sum()}\n" +
                "\t����������� ������: ${it.tests}\t�����: ${it.tests.sum() * 2}\n" +
                "����� ������: ${it.getTotalScore() + it.ex + it.exams.sum()}\n")
    }
    /**
     * 3 ������������ ������� ��� ��, ��� � � ������� ������� ����������. ������� 5
    ������ ����� ���������� ������, ������������ �� �� �������. ������� 5 ������ ����� ��
    ���������� ������, ������������ �� �� ������� � ��� �� �������, ��� � � ������� �������,
    �� ���� �������� �������� ����������.
     */
    val rating = students.sortedBy { it.name }.sortedByDescending { it.getTotalScore() + it.ex }

    val studOnlyAccess = rating.filter { it.evalPreGrade() == "������" }.takeLast(5)
    val studNoAccess = rating.filter { it.evalPreGrade() == "��������" }.take(5)

    println("�  ${"������".padEnd(6)} ${"������� � �".padEnd(20)} ${"��".padEnd(2)} " +
            "${"�������".padEnd(8)} ${"���".padEnd(3)} " +
            "${"��".padEnd(5)} " +
            "${"��".padEnd(5)} " +
            "${"��".padEnd(5)} " +
            "${"������".padEnd(6)} " +
            "${"�����".padEnd(5)}")
    println()
    studOnlyAccess.forEachIndexed { index, it -> println("${(index + 1).toString().padEnd(2)} ${it.group.padEnd(6)} ${it.name.padEnd(20)} ${it.attest.padEnd(2)} " +
            "${it.evalPreGrade().padEnd(8)} ${(it.attendanceGrade() * 10).toString().padEnd(3)} " +
            "${String.format("%.1f", it.labWorks.sum().toDouble() / (it.labWorks.size * 5) * 100).padEnd(5)} " +
            "${String.format("%.1f", it.exams.sum().toDouble() / (it.exams.size * 5) * 100).padEnd(5)} " +
            "${String.format("%.1f", it.tests.sum().toDouble() / (it.tests.size * 5) * 100).padEnd(5)} " +
            "${String.format("%.1f", it.getTotalScore().toDouble() / (it.attendanceGrade() + it.labWorks.size * 5 + it.tests.size * 10) * 100).padEnd(6)} " +
            "${(it.getTotalScore() + it.ex + it.exams.sum()).toString().padEnd(5)}") }
    println()
    studNoAccess.forEachIndexed { index, it -> println("${(index + 1).toString().padEnd(2)} ${it.group.padEnd(6)} ${it.name.padEnd(20)} ${it.attest.padEnd(2)} " +
            "${it.evalPreGrade().padEnd(10)} ${(it.attendanceGrade() * 10).toString().padEnd(3)} " +
            "${String.format("%.1f", it.labWorks.sum().toDouble() / (it.labWorks.size * 5) * 100).padEnd(5)} " +
            "${String.format("%.1f", it.exams.sum().toDouble() / (it.exams.size * 5) * 100).padEnd(5)} " +
            "${String.format("%.1f", it.tests.sum().toDouble() / (it.tests.size * 5) * 100).padEnd(5)} " +
            "${String.format("%.1f", it.getTotalScore().toDouble() / (it.attendanceGrade() + it.labWorks.size * 5 + it.tests.size * 10) * 100).padEnd(6)} " +
            "${(it.getTotalScore() + it.ex + it.exams.sum()).toString().padEnd(5)}") }

    /**
     * 4 ������������ ���������� ������ ������ ��� ��, ��� ������������ � ������� �
    ������� �� �����.
     */
    val groupByGroup = students.groupBy { it.group }
    val results = groupByGroup.mapValues { entry ->
        val groupStudents = entry.value

        val avgAttendance = groupStudents.map { it.attendanceGrade() }.average() // ������� ������������
        val avgLabWorks = averageIntList(groupStudents.map { it.labWorks }) // ������� ���� �� ��������� ��
        val avgExams = averageIntList(groupStudents.map { it.exams }) // ������� ���� �� ��������� ��
        val avgTests = averageIntList(groupStudents.map { it.tests }) // ������� ���� �� ��������� ��
        val commonAttest = mostCommonAttest(groupStudents.map { it.attest }) // ���������� �� ������

        val avgCommonLabWorks = groupStudents.map { it.labWorks.sum() }.average() // ������� ���� �� ��� ��
        val avgCommonExams = groupStudents.map { it.exams.sum() }.average() // ������� ���� �� ��� ��
        val avgCommonTests = groupStudents.map { it.tests.sum() * 2 }.average() // ������� ���� �� ��� ��
        val avgCommonAccess = groupStudents.map { it.attendanceGrade() + it.labWorks.sum() + it.tests.sum() * 2 }.average() // ������� ���� �� ������
        val avgCommonTotal = groupStudents.map { it.attendanceGrade() + it.labWorks.sum() + it.tests.sum() * 2 + it.ex}.average() // ������� ���� �����

        val avgPercentLabWorks = groupStudents.map { it.labWorks.sum().toDouble() / (it.labWorks.size * 5) * 100 }.average() // ������� ��
        val avgPercentExams = groupStudents.map { it.exams.sum().toDouble() / (it.exams.size * 5) * 100 }.average() // ������� ��
        val avgPercentTests = groupStudents.map { it.tests.sum().toDouble() * 2 / (it.tests.size * 10) * 100 }.average() // ������� ��
        val avgPercentAccess = groupStudents.map { (it.attendanceGrade() + it.labWorks.sum() + it.tests.sum() * 2).toDouble() / 75 * 100 }.average() // ������� ������
        val avgPercentTotal = groupStudents.map { it.attendanceGrade() + it.labWorks.sum() + it.tests.sum() * 2 + it.ex }.average() // ������� �����

        GroupStats(avgAttendance, avgLabWorks, avgExams, avgTests, commonAttest,
            avgCommonLabWorks, avgCommonExams, avgCommonTests, avgCommonAccess, avgCommonTotal,
            avgPercentLabWorks, avgPercentExams, avgPercentTests, avgPercentAccess, avgPercentTotal)
    }

    // ������� �����
    println("${"������".padEnd(6)} ${"��1".padEnd(3)} ${"��2".padEnd(3)} ${"��3".padEnd(3)} " +
            "${"��4".padEnd(3)} ${"��5".padEnd(3)} ${"��6".padEnd(3)} ${"��7".padEnd(3)} " +
            "${"��1".padEnd(3)} ${"��2".padEnd(3)} ${"��3".padEnd(3)} ${"��4".padEnd(3)} ${"��5".padEnd(3)} " +
            "${"��1".padEnd(3)} ${"��2".padEnd(3)} ${"��3".padEnd(3)} " +
            "${"���".padEnd(4)} ${"��".padEnd(4)} ${"��".padEnd(4)} ${"��".padEnd(4)} " +
            "${"������".padEnd(6)} ${"�����".padEnd(5)} ${"��".padEnd(2)} ${"������".padEnd(8)} " +
            "${"���".padEnd(4)} ${"��".padEnd(4)} ${"��".padEnd(4)} ${"��".padEnd(4)} " +
            "${"������".padEnd(6)} ${"�����".padEnd(5)}")
    results.forEach { (group, stats) ->
        print("${group.padEnd(6)} ")
        for (i in 0..6)
            print(printFormatted(stats.avgLabWorks[i]))
        for (i in 0..4)
            print(printFormatted(stats.avgExams[i]))
        for (i in 0..2)
            print(printFormatted(stats.avgTests[i]))
        print("${String.format("%.1f", stats.avgAttendance).padEnd(4)} " +
                "${String.format("%.1f", stats.avgCommonLabWorks).padEnd(4)} " +
                "${String.format("%.1f", stats.avgCommonExams).padEnd(4)} " +
                "${String.format("%.1f", stats.avgCommonTests).padEnd(4)} " +
                "${String.format("%.1f", stats.avgCommonAccess).padEnd(6)} " +
                "${String.format("%.1f", stats.avgCommonTotal).padEnd(5)} " +
                "${stats.commonAttest.padEnd(2)} ")
        print((when {
            stats.avgCommonAccess < 30 -> "��������"
            stats.avgCommonTotal < 60 -> "������"
            stats.avgCommonTotal < 75 -> "�������"
            stats.avgCommonTotal < 90 -> "������"
            else -> "�������"
        }).padEnd(8))
        print(" ${(String.format("%.0f", stats.avgAttendance * 10) + "%").padEnd(4)} " +
                "${(String.format("%.0f", stats.avgPercentLabWorks) + "%").padEnd(4)} " +
                "${(String.format("%.0f", stats.avgPercentExams) + "%").padEnd(4)} " +
                "${(String.format("%.0f", stats.avgPercentTests) + "%").padEnd(4)} " +
                "${(String.format("%.0f", stats.avgPercentAccess) + "%").padEnd(6)} " +
                "${(String.format("%.0f", stats.avgPercentTotal) + "%").padEnd(4)} ")
        println()
    }
    // ������� �������� �� ������� �������
    print("3 ���� ")
    for (i in 0..6)
        print(printFormatted(results.map { (_, stats) -> stats.avgLabWorks[i] }.average()))
    for (i in 0..4)
        print(printFormatted(results.map { (_, stats) -> stats.avgExams[i] }.average()))
    for (i in 0..2)
        print(printFormatted(results.map { (_, stats) -> stats.avgTests[i] }.average()))
    print("${String.format("%.1f", results.map { (_, stats) -> stats.avgAttendance }.average()).padEnd(4)} " +
            "${String.format("%.1f", results.map { (_, stats) -> stats.avgCommonLabWorks }.average()).padEnd(4)} " +
            "${String.format("%.1f", results.map { (_, stats) -> stats.avgCommonExams }.average()).padEnd(4)} " +
            "${String.format("%.1f", results.map { (_, stats) -> stats.avgCommonTests }.average()).padEnd(4)} " +
            "${String.format("%.1f", results.map { (_, stats) -> stats.avgCommonAccess }.average()).padEnd(6)} " +
            "${String.format("%.1f", results.map { (_, stats) -> stats.avgCommonTotal }.average()).padEnd(5)} " +
            "�� ")
    print((when {
        results.map { (_, stats) -> stats.avgCommonAccess }.average() <= 75 * 0.4 -> "��������"
        results.map { (_, stats) -> stats.avgCommonTotal }.average() < 60 -> "������"
        results.map { (_, stats) -> stats.avgCommonTotal }.average() < 75 -> "�������"
        results.map { (_, stats) -> stats.avgCommonTotal }.average() < 90 -> "������"
        else -> "�������"
    }).padEnd(8))
    print(" ${(String.format("%.0f", results.map { (_, stats) -> stats.avgAttendance * 10 }.average()) + "%").padEnd(4)} " +
            "${(String.format("%.0f", results.map { (_, stats) -> stats.avgPercentLabWorks }.average()) + "%").padEnd(4)} " +
            "${(String.format("%.0f", results.map { (_, stats) -> stats.avgPercentExams }.average()) + "%").padEnd(4)} " +
            "${(String.format("%.0f", results.map { (_, stats) -> stats.avgPercentTests }.average()) + "%").padEnd(4)} " +
            "${(String.format("%.0f", results.map { (_, stats) -> stats.avgPercentAccess }.average()) + "%").padEnd(6)} " +
            "${(String.format("%.0f", results.map { (_, stats) -> stats.avgPercentTotal }.average()) + "%").padEnd(5)} ")
    /**
     * 5 ����� ������, � ������� ���������� ������� ���� �� ������ ����� ���, ��� �������
    ������ 3, 4, 5
     */
    val autoExam = students.filter { it.evalPreGrade() in listOf("�������", "������", "�������") }
    val groupByGroupWithEx = autoExam.groupBy { it.group }
    val avgInGroup = groupByGroupWithEx.mapValues { (_, group) -> group.map { it.getTotalScore().toDouble() / (it.attendanceGrade() + it.labWorks.size * 5 + it.tests.size * 10) * 100 }.average() }
    println("� ������ ${avgInGroup.minByOrNull { it.value }?.key} ���������� ������� ���� �� ������ ����� ���, ��� ������� 3, 4 ��� 5")
}
