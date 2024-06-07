import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File

class Student(val group: String, // группа
              val name: String, // ФИО
              val ex: Int, // доп баллы
              val lecAttendance: List<Boolean>, // посещаемость лекций
              val labAttendance: List<Boolean>, // посещаемость лаб
              val labWorks: List<Int>, // оценки за ЛР
              val exams: List<Int>, // оценки за ЭП
              val tests: List<Int>, // оценки за КР
              val attest: String) { // аттестация
// Вывод студента
    fun write() {
        println("Студент: $name\n" +
                "Доп баллы: $ex\n" +
                "Посещаемость: ${attendanceGrade()}\n" +
                "Оценки за:\n" +
                "\tЛабораторные работы: $labWorks\n" +
                "\tЭкзаменационный проект: $exams\n" +
                "\tКонтрольные работы: $tests\n" +
                "Итого баллов: ${getTotalScore() + ex}\n")
    }
// Подсчёт оценки посещаемости
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
// Подсчёт всех баллов для допуска (без ЭП и доп баллов)
    fun getTotalScore(): Int {
        return attendanceGrade() + labWorks.sum() + tests.sum() * 2
    }
// Получение данных об экзамене
    fun evalPreGrade(): String {
        return when {
            getTotalScore() < 30 -> "недопуск"
            (getTotalScore() + ex) < 60 -> "допуск"
            (getTotalScore() + ex) < 75 -> "САМОКАТ"
            (getTotalScore() + ex) < 90 -> "ХОРОШО"
            else -> "ОТЛИЧНО"
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
// Список средних значений из Int
fun averageIntList(lists: List<List<Int>>): List<Double> {
    if (lists.isEmpty()) return emptyList()

    val size = lists.first().size
    return List(size) { index ->
        lists.map { it[index] }.average()
    }
}

// Выбор ат ил н для группы в зависимости от того, чего больше
fun mostCommonAttest(attests: List<String>): String {
    return attests.groupBy { it }.maxByOrNull { it.value.size }?.key ?: ""
}

// Получение оценки из ячейки. Если там число, то возвращаем его, если строка, то преобразуем в баллы
fun getCellGrade(cell: Cell): Int{
    return when (cell.cellType) {
        CellType.STRING -> {
            val stringValue = cell.stringCellValue
            when (stringValue) {
                "д" -> 2
                "", "-" -> 0
                else -> stringValue.toIntOrNull()
                    ?: 0 // Попробуем преобразовать строку в число, если это не удастся, то 0
            }
        }
        else -> cell.numericCellValue.toInt()
    }
}

// Чтение файла
fun fileRead(): List<Student> {
    val students = mutableListOf<Student>()

    val file = File("students.xlsx")
    val workbook = WorkbookFactory.create(file.inputStream())
    // Проход по всем листам
    for (sheetNum in 0 until workbook.numberOfSheets) {
        val sheet = workbook.getSheetAt(sheetNum)
        var rowC = 0
        // Выбор листов групп
        if (("36" in sheet.sheetName) or ("39" in sheet.sheetName))
            for (row in sheet) {
                // Пропуск двух строк
                if (rowC > 1)
                    if (row.getCell(0) != null) {
                        val name = row.getCell(1).stringCellValue // получение ФИО
                        val ex = row.getCell(2).numericCellValue.toInt() // получение доп баллов

                        val lecAttendanceList = mutableListOf<Boolean>()
                        val labAttendanceList = mutableListOf<Boolean>()
                        var cellIndex = 4
                        // Чтение данных о посещаемости отдельно для лекций и для лаб
                        var leclab = true
                        while (sheet.getRow(0).getCell(cellIndex).cellType != CellType.STRING) {
                            // переключение счёта лекций и лаб
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

                        // Чтение оценок за лабораторные работы
                        val labWorks = mutableListOf<Int>()
                        for (i in cellIndex..(cellIndex + 6)) { // Колонки с ЛР1 по ЛР7
                            val value = getCellGrade(row.getCell(i))
                            labWorks.add(value)
                        }
                        cellIndex += 7

                        // Чтение оценок за экзаменационные работы
                        val exams = mutableListOf<Int>()
                        for (i in cellIndex..(cellIndex + 4)) { // Колонки с ЭП1 по ЭП5
                            val value = getCellGrade(row.getCell(i))
                            exams.add(value)
                        }
                        cellIndex += 5

                        // Чтение оценок за контрольные работы
                        val tests = mutableListOf<Int>()
                        for (i in cellIndex..(cellIndex + 2)) { // Колонки с КР1 по КР3
                            val value = getCellGrade(row.getCell(i))
                            tests.add(value)
                        }
                        cellIndex += 32

                        // Создание объекта Student и добавление его в список
                        val student = Student(sheet.sheetName, name, ex, lecAttendanceList, labAttendanceList, labWorks, exams, tests, row.getCell(cellIndex).stringCellValue)
                        students.add(student)
                    } else break
                rowC += 1
            }
    }

    workbook.close()

    return students
}

// Вывод в таблицу групп значений ЛР, ЭП и КР
fun printFormatted(value: Double): String {
    return when {
        value < 0.5 -> "-".padEnd(4)
        value < 1 -> "+".padEnd(4)
        value < 1.5 -> "++".padEnd(4)
        value < 2.5 -> "д".padEnd(4)
        else -> "${String.format("%.1f", value).padEnd(3)} "
    }
}

fun main() {
    /**
     * Дана таблица с результатами обучения Вашей группы в семестре. Необходимо
    прочитать все данные о посещаемости и все оценки за КР, ЛР и ЭП. Создать структуру классов
    для хранения информации.
    1 Прочитать суммарные баллы и оценки. Найти лабораторную работу, которую
    выполнили наибольшее количество человек на любую оценку среди неполучивших допуск.
    Найти КР, которую выполнили наименьшее количество человек хотя бы на допуск среди не
    поливших допуск.
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
    println("Лабораторная работа №$maxL была сдана наибольшим количеством студентов без допуска")
    println("Контрольная работа №$minK была сдана наименьшим количеством студентов без допуска\n")

    /**
     * 2 Рассчитать самостоятельно в программе баллы за посещаемость, КР, ЛР и ЭП так же,
    как в документе. Найти сумму баллов и вывести количество баллов за посещаемость, ЛР, ЭП,
    КР, ИТОГО для каждого студента.
     */
    students.forEach {
        println("Студент: ${it.name}\n" +
                "Доп баллы: ${it.ex}\n" +
                "Посещаемость: ${it.attendanceGrade()}\n" +
                "Оценки за:\n" +
                "\tЛабораторные работы: ${it.labWorks}\tВсего: ${it.labWorks.sum()}\n" +
                "\tЭкзаменационный проект: ${it.exams}\tВсего: ${it.exams.sum()}\n" +
                "\tКонтрольные работы: ${it.tests}\tВсего: ${it.tests.sum() * 2}\n" +
                "Итого баллов: ${it.getTotalScore() + it.ex + it.exams.sum()}\n")
    }
    /**
     * 3 Сформировать рейтинг так же, как и в разделе рейтинг программно. Вывести 5
    худших среди получивших допуск, сгруппировав их по группам. Вывести 5 лучших среди не
    получивших допуск, сгруппировав их по группам в том же формате, как и в разделе рейтинг,
    то есть выводить проценты сделанного.
     */
    val rating = students.sortedBy { it.name }.sortedByDescending { it.getTotalScore() + it.ex }

    val studOnlyAccess = rating.filter { it.evalPreGrade() == "допуск" }.takeLast(5)
    val studNoAccess = rating.filter { it.evalPreGrade() == "недопуск" }.take(5)

    println("№  ${"Группа".padEnd(6)} ${"Фамилия И О".padEnd(20)} ${"ат".padEnd(2)} " +
            "${"Экзамен".padEnd(8)} ${"Пос".padEnd(3)} " +
            "${"ЛР".padEnd(5)} " +
            "${"ЭП".padEnd(5)} " +
            "${"КР".padEnd(5)} " +
            "${"Допуск".padEnd(6)} " +
            "${"Итого".padEnd(5)}")
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
     * 4 Сформировать программно раздел группы так же, как сформировано в таблице и
    вывести на экран.
     */
    val groupByGroup = students.groupBy { it.group }
    val results = groupByGroup.mapValues { entry ->
        val groupStudents = entry.value

        val avgAttendance = groupStudents.map { it.attendanceGrade() }.average() // средняя посещаемость
        val avgLabWorks = averageIntList(groupStudents.map { it.labWorks }) // средний балл за отдельные ЛР
        val avgExams = averageIntList(groupStudents.map { it.exams }) // средний балл за отдельные ЭП
        val avgTests = averageIntList(groupStudents.map { it.tests }) // средний балл за отдельные КР
        val commonAttest = mostCommonAttest(groupStudents.map { it.attest }) // аттестация по группе

        val avgCommonLabWorks = groupStudents.map { it.labWorks.sum() }.average() // средний балл за все ЛР
        val avgCommonExams = groupStudents.map { it.exams.sum() }.average() // средний балл за все ЭП
        val avgCommonTests = groupStudents.map { it.tests.sum() * 2 }.average() // средний балл за все КР
        val avgCommonAccess = groupStudents.map { it.attendanceGrade() + it.labWorks.sum() + it.tests.sum() * 2 }.average() // средний балл на допуск
        val avgCommonTotal = groupStudents.map { it.attendanceGrade() + it.labWorks.sum() + it.tests.sum() * 2 + it.ex}.average() // средний балл итого

        val avgPercentLabWorks = groupStudents.map { it.labWorks.sum().toDouble() / (it.labWorks.size * 5) * 100 }.average() // процент ЛР
        val avgPercentExams = groupStudents.map { it.exams.sum().toDouble() / (it.exams.size * 5) * 100 }.average() // процент ЭП
        val avgPercentTests = groupStudents.map { it.tests.sum().toDouble() * 2 / (it.tests.size * 10) * 100 }.average() // процент КР
        val avgPercentAccess = groupStudents.map { (it.attendanceGrade() + it.labWorks.sum() + it.tests.sum() * 2).toDouble() / 75 * 100 }.average() // процент допуск
        val avgPercentTotal = groupStudents.map { it.attendanceGrade() + it.labWorks.sum() + it.tests.sum() * 2 + it.ex }.average() // процент итого

        GroupStats(avgAttendance, avgLabWorks, avgExams, avgTests, commonAttest,
            avgCommonLabWorks, avgCommonExams, avgCommonTests, avgCommonAccess, avgCommonTotal,
            avgPercentLabWorks, avgPercentExams, avgPercentTests, avgPercentAccess, avgPercentTotal)
    }

    // Таблица групп
    println("${"Группа".padEnd(6)} ${"ЛР1".padEnd(3)} ${"ЛР2".padEnd(3)} ${"ЛР3".padEnd(3)} " +
            "${"ЛР4".padEnd(3)} ${"ЛР5".padEnd(3)} ${"ЛР6".padEnd(3)} ${"ЛР7".padEnd(3)} " +
            "${"ЭП1".padEnd(3)} ${"ЭП2".padEnd(3)} ${"ЭП3".padEnd(3)} ${"ЭП4".padEnd(3)} ${"ЭП5".padEnd(3)} " +
            "${"КР1".padEnd(3)} ${"КР2".padEnd(3)} ${"КР3".padEnd(3)} " +
            "${"Пос".padEnd(4)} ${"ЛР".padEnd(4)} ${"ЭП".padEnd(4)} ${"КР".padEnd(4)} " +
            "${"Допуск".padEnd(6)} ${"Итого".padEnd(5)} ${"Ат".padEnd(2)} ${"Оценка".padEnd(8)} " +
            "${"Пос".padEnd(4)} ${"ЛР".padEnd(4)} ${"ЭП".padEnd(4)} ${"КР".padEnd(4)} " +
            "${"Допуск".padEnd(6)} ${"Итого".padEnd(5)}")
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
            stats.avgCommonAccess < 30 -> "недопуск"
            stats.avgCommonTotal < 60 -> "допуск"
            stats.avgCommonTotal < 75 -> "САМОКАТ"
            stats.avgCommonTotal < 90 -> "ХОРОШО"
            else -> "ОТЛИЧНО"
        }).padEnd(8))
        print(" ${(String.format("%.0f", stats.avgAttendance * 10) + "%").padEnd(4)} " +
                "${(String.format("%.0f", stats.avgPercentLabWorks) + "%").padEnd(4)} " +
                "${(String.format("%.0f", stats.avgPercentExams) + "%").padEnd(4)} " +
                "${(String.format("%.0f", stats.avgPercentTests) + "%").padEnd(4)} " +
                "${(String.format("%.0f", stats.avgPercentAccess) + "%").padEnd(6)} " +
                "${(String.format("%.0f", stats.avgPercentTotal) + "%").padEnd(4)} ")
        println()
    }
    // Средние значения по каждому столбцу
    print("3 курс ")
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
            "ат ")
    print((when {
        results.map { (_, stats) -> stats.avgCommonAccess }.average() <= 75 * 0.4 -> "недопуск"
        results.map { (_, stats) -> stats.avgCommonTotal }.average() < 60 -> "допуск"
        results.map { (_, stats) -> stats.avgCommonTotal }.average() < 75 -> "САМОКАТ"
        results.map { (_, stats) -> stats.avgCommonTotal }.average() < 90 -> "ХОРОШО"
        else -> "ОТЛИЧНО"
    }).padEnd(8))
    print(" ${(String.format("%.0f", results.map { (_, stats) -> stats.avgAttendance * 10 }.average()) + "%").padEnd(4)} " +
            "${(String.format("%.0f", results.map { (_, stats) -> stats.avgPercentLabWorks }.average()) + "%").padEnd(4)} " +
            "${(String.format("%.0f", results.map { (_, stats) -> stats.avgPercentExams }.average()) + "%").padEnd(4)} " +
            "${(String.format("%.0f", results.map { (_, stats) -> stats.avgPercentTests }.average()) + "%").padEnd(4)} " +
            "${(String.format("%.0f", results.map { (_, stats) -> stats.avgPercentAccess }.average()) + "%").padEnd(6)} " +
            "${(String.format("%.0f", results.map { (_, stats) -> stats.avgPercentTotal }.average()) + "%").padEnd(5)} ")
    /**
     * 5 Найти группу, в которой наименьший средний балл на допуск среди тех, кто получил
    оценки 3, 4, 5
     */
    val autoExam = students.filter { it.evalPreGrade() in listOf("ОТЛИЧНО", "ХОРОШО", "САМОКАТ") }
    val groupByGroupWithEx = autoExam.groupBy { it.group }
    val avgInGroup = groupByGroupWithEx.mapValues { (_, group) -> group.map { it.getTotalScore().toDouble() / (it.attendanceGrade() + it.labWorks.size * 5 + it.tests.size * 10) * 100 }.average() }
    println("В группе ${avgInGroup.minByOrNull { it.value }?.key} наименьший средний балл на допуск среди тех, кто получил 3, 4 или 5")
}
