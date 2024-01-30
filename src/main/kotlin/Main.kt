import org.dhatim.fastexcel.reader.Cell
import org.dhatim.fastexcel.reader.ReadableWorkbook
import org.dhatim.fastexcel.reader.Row
import java.io.FileInputStream
import java.util.stream.Collectors
import java.util.stream.Stream

fun main(args: Array<String>) {
    val start = System.currentTimeMillis()
    FileInputStream("*****.xlsx").use { fis ->
        val wb = ReadableWorkbook(fis)
        val cellValues = wb.firstSheet.openStream()
            .parallel()
            .skip(1)
            .map { row: Row ->
                rowCellsToDataClass(
                    row.stream()
                )
            }
            .collect(Collectors.toList())
        val list = cellValues.map { cellValue ->
            cellValue.zipCode
        }.toSet().sortedBy { it -> Integer.parseInt(it) }


        println("cellValues size: ${cellValues.size}")
        println("list size: ${list.size}")
        println("list: $list")
        println("time: ${System.currentTimeMillis() - start}")
    }
}

private fun rowCellsToDataClass(stream: Stream<Cell>): CellValue {
    val cells = stream.collect(Collectors.toList())
    return CellValue(cells[0].value.toString())
}

data class CellValue(var zipCode: String, var secondName: String?) {
    constructor(firstName: String) : this(firstName, null)
}
