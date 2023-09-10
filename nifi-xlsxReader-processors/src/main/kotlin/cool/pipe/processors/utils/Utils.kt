package cool.pipe.processors.utils

import com.fasterxml.jackson.databind.ObjectMapper
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.FormulaEvaluator
import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.time.format.DateTimeFormatter

/**
 * Reads data from an Excel (.xlsx) file and converts it into a JSON array.
 *
 * @param path The path to the Excel file to read.
 * @param range The cell range to read, specified as "startCell:endCell" (e.g., "A1:B10").
 * @param sheetIndex The index of the sheet to read (default is 0 for the first sheet).
 * @param headers Specify whether to include headers in the JSON output (default is true).
 * @param formatDateHeader The date format for formatting header date cells (default is "yyyy-MM-dd").
 * @param formatDateBody The date format for formatting date cells within the data (default is "yyyy-MM-dd HH:mm:ss.SSS").
 * @return A JSON-formatted string containing the data from the specified Excel range.
 */
fun readXlsx(
    path: String,
    range: String = "A1:A1",
    sheetIndex: Int = 0,
    headers: Boolean = true,
    formatDateHeader: String = "yyyy-MM-dd",
    formatDateBody: String = "yyyy-MM-dd HH:mm:ss.SSS"
): String {

    val rangeSplit = range.split(":")
    val colStart = rangeSplit[0].replace("[^A-Za-z]+".toRegex(), "")
    val colStartIndex = CellReference.convertColStringToIndex(colStart)
    val colEnd = rangeSplit[1].replace("[^A-Za-z]+".toRegex(), "")
    val colEndIndex = CellReference.convertColStringToIndex(colEnd)
    var rowStart = rangeSplit[0].replace("\\D+".toRegex(), "").toInt()
    val rowEnd = rangeSplit[1].replace("\\D+".toRegex(), "").toInt()

    //formato de fecha de salida
    val formatoDateHeader = DateTimeFormatter.ofPattern(formatDateHeader)
    val formatoDateBody = DateTimeFormatter.ofPattern(formatDateBody)

    try {
        val file = FileInputStream(File(path))

        //Create Workbook instance holding reference to .xlsx file
        val workbook = XSSFWorkbook(file)

        val evaluator: FormulaEvaluator = workbook.creationHelper.createFormulaEvaluator()

        //Get first/desired sheet from the workbook
        val sheet = workbook.getSheetAt(sheetIndex)

        //generamos los headers
        val headersList = mutableListOf<String>()
        var notFound = 0

        //si pedimos headers, partimos desde la siguiente
        if (headers) {

            val headersRow = sheet.getRow(rowStart - 1)
            for (column in colStartIndex..colEndIndex) {

                val cellTemp = headersRow.getCell(column)
                if (cellTemp != null && cellTemp.cellType != CellType.BLANK && cellTemp.cellType != CellType.ERROR) {

                    when (cellTemp.cellType) {
                        CellType.BOOLEAN -> {
                            headersList.add(cellTemp.booleanCellValue.toString())
                        }

                        CellType.NUMERIC -> {
                            if (DateUtil.isCellDateFormatted(cellTemp)) {
                                val formattedDate = cellTemp.localDateTimeCellValue.format(formatoDateHeader)
                                headersList.add(formattedDate)
                            } else {
                                headersList.add(cellTemp.numericCellValue.toString())
                            }
                        }

                        CellType.STRING -> {
                            headersList.add(cellTemp.stringCellValue)
                        }

                        CellType.FORMULA -> {

                            when (evaluator.evaluate(cellTemp).cellType) {
                                CellType.BOOLEAN -> {
                                    headersList.add(evaluator.evaluate(cellTemp).booleanValue.toString())
                                }

                                CellType.NUMERIC -> {
                                    if (DateUtil.isCellDateFormatted(evaluator.evaluateInCell(cellTemp))) {
                                        val formattedDate =
                                            evaluator.evaluateInCell(cellTemp).localDateTimeCellValue.format(
                                                formatoDateHeader
                                            )
                                        headersList.add(formattedDate)
                                    } else {
                                        headersList.add(evaluator.evaluate(cellTemp).numberValue.toString())
                                    }
                                }

                                CellType.STRING -> {
                                    headersList.add(evaluator.evaluate(cellTemp).stringValue)
                                }

                                CellType.ERROR -> {
                                    headersList.add(evaluator.evaluate(cellTemp).errorValue.toString())
                                }

                                CellType.BLANK -> {
                                    headersList.add(notFound.toString())
                                }

                                else -> {
                                    headersList.add(notFound.toString())
                                }
                            }


                        }

                        else -> {
                            headersList.add(notFound.toString())
                        }
                    }
                } else {
                    headersList.add(notFound.toString())
                }
                notFound++
            }

            //actualiza el valor de rowStart para considerar el header
            rowStart += 1
        } else {
            //generamos los headers
            for ((index, _) in (colStartIndex..colEndIndex).withIndex()) {
                headersList.add(index.toString())
            }
        }

        val objectMapper = ObjectMapper()

        //creamos array para json
        val array = objectMapper.createArrayNode()

        for (row in rowStart..rowEnd) {

            //creamos objeto para el array
            val objectJson = objectMapper.createObjectNode()

            val tempRow = sheet.getRow(row - 1)
            for ((index, column) in (colStartIndex..colEndIndex).withIndex()) {

                if (tempRow != null) {
                    val cellTemp = tempRow.getCell(column)

                    if (cellTemp != null) {
                        when (cellTemp.cellType) {
                            CellType.BOOLEAN -> {
                                objectJson.put(headersList[index], cellTemp.booleanCellValue)
                            }

                            CellType.NUMERIC -> {
                                if (DateUtil.isCellDateFormatted(cellTemp)) {
                                    val formattedDate = cellTemp.localDateTimeCellValue.format(formatoDateBody)
                                    objectJson.put(headersList[index], formattedDate)
                                } else {
                                    objectJson.put(headersList[index], cellTemp.numericCellValue)
                                }
                            }

                            CellType.STRING -> {
                                objectJson.put(headersList[index], cellTemp.stringCellValue)
                            }

                            CellType.FORMULA -> {

                                when (evaluator.evaluate(cellTemp).cellType) {
                                    CellType.BOOLEAN -> {
                                        objectJson.put(headersList[index], evaluator.evaluate(cellTemp).booleanValue)
                                    }

                                    CellType.NUMERIC -> {
                                        if (DateUtil.isCellDateFormatted(evaluator.evaluateInCell(cellTemp))) {
                                            val formattedDate =
                                                evaluator.evaluateInCell(cellTemp).localDateTimeCellValue.format(
                                                    formatoDateHeader
                                                )
                                            objectJson.put(headersList[index], formattedDate)
                                        } else {
                                            objectJson.put(headersList[index], evaluator.evaluate(cellTemp).numberValue)
                                        }
                                    }

                                    CellType.STRING -> {
                                        objectJson.put(headersList[index], evaluator.evaluate(cellTemp).stringValue)
                                    }

                                    CellType.ERROR -> {
                                        objectJson.put(
                                            headersList[index],
                                            evaluator.evaluate(cellTemp).errorValue.toString()
                                        )
                                    }

                                    CellType.BLANK -> {
                                        objectJson.put(headersList[index], "")
                                    }

                                    else -> {
                                        objectJson.put(headersList[index], "")
                                    }
                                }

                            }

                            CellType.BLANK -> {
                                objectJson.put(headersList[index], "")
                            }

                            CellType.ERROR -> {
                                objectJson.put(headersList[index], cellTemp.errorCellString)
                            }

                            else -> {
                                objectJson.put(headersList[index], "")
                            }
                        }
                    } else {
                        objectJson.put(headersList[index], "")
                    }
                } else {
                    objectJson.put(headersList[index], "")
                }
            }
            array.add(objectJson)
        }

        workbook.close()
        file.close()

        return array.toPrettyString()
    } catch (e: Exception) {
        e.printStackTrace()
    }
    return ""
}