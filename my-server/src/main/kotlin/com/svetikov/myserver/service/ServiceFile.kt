package com.svetikov.myserver.service

import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.tomcat.util.http.fileupload.FileUploadException
import org.springframework.beans.factory.annotation.Value
import org.springframework.stereotype.Service
import org.springframework.web.multipart.MultipartFile
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream

import java.nio.file.Files
import java.nio.file.Paths
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter
import java.util.concurrent.TimeUnit
import java.util.logging.Logger
import kotlin.Exception
import kotlin.io.path.deleteIfExists
import kotlin.math.pow
import kotlin.math.roundToInt
import kotlin.math.sqrt

@Service
class ServiceFile {
    val logServiceFile = Logger.getLogger(ServiceFile::class.java.name)
    private var nameFileIMAL: String? = "MDF TEST RAPORT1.xls"
    private var nameFilePLC: String? = "dataTest2021 09 20 07 18 22.xlsx"

    private var path =""

    //todo init list from imal
    private val listDicke = mutableListOf<String>()
    private val listRohdichte = mutableListOf<String>()
    private val listQuerzug = mutableListOf<String>()
    private val listAbhebe = mutableListOf<String>()
    private val listBiege = mutableListOf<String>()
    private val listEModul = mutableListOf<String>()
    private val listQuellung24h = mutableListOf<String>()
    private val listRestfeuchte = mutableListOf<String>()
    private val listWasseraufnahme = mutableListOf<String>()

    //todo init list from plc
    private val listDryerPlc = mutableListOf<String>()
    private val listRefinerPlc = mutableListOf<String>()
    private val listGlueKitchenPlc = mutableListOf<String>()
    private val listDosingChipsPlc = mutableListOf<String>()
    private val listMatformerPresssPlc = mutableListOf<String>()
    private val listProductInfoPlc = mutableListOf<String>()


    @Value("\${upload.path}")
    lateinit var uploadPath: String//my-server/src/main/resources/upload

    @Value("\${upload_plc.path_plc}")
    lateinit var uploadPathPlc: String//my-server/src/main/resources/upload

    fun save(file: MultipartFile) {
        try {
            val root = Paths.get(uploadPath)
            val resolve = root.resolve(file.originalFilename)
            println(file.originalFilename)

            nameFileIMAL = file.originalFilename
            logServiceFile.info("save() nameFileIMAL $nameFileIMAL")
            if (resolve.toFile().exists()) {
                val status = resolve.deleteIfExists()
                if (!status)
                    throw FileUploadException("File already exists ${file.originalFilename}")
            }
            Files.copy(file.inputStream, resolve)
        } catch (e: Exception) {
            throw FileUploadException("Couldn't store the file.Error ${e.message}")
        }
    }

    //todo upload file from PLC SIMATIC S7 dataTest2021 09 20 07 18 22
    fun savePlc(filePlc: MultipartFile) {
        try {
            val rootPlc = Paths.get(uploadPathPlc)
            val resolvePlc = rootPlc.resolve(filePlc.originalFilename)
            println(filePlc.originalFilename)

            nameFilePLC = filePlc.originalFilename
            logServiceFile.info("save() nameFilePlc $nameFilePLC")
            if (resolvePlc.toFile().exists()) {
                val statusPlc = resolvePlc.deleteIfExists()
                if (!statusPlc)
                    throw FileUploadException("File already exists ${filePlc.originalFilename}")
            }
            Files.copy(filePlc.inputStream, resolvePlc)
        } catch (e: Exception) {
            throw FileUploadException("Couldn't store the file.Error ${e.message}")
        }
    }

    fun readIMALAndCreateTableLaboratoryReport() {
        logServiceFile.info("readIMALAndCreateTableLaboratoryReport() ")
        readTableFromIMAL(
            nameFileIMAL,
            "Dicke" to listDicke,
            "Rohdichte" to listRohdichte,
            "Querzug" to listQuerzug,
            "Abhebe" to listAbhebe,
            "Biege" to listBiege,
            "EModul" to listEModul,
            "Quellung24h" to listQuellung24h,
            "Restfeuchte" to listRestfeuchte,
            "Wasseraufnahme" to listWasseraufnahme

        )

        logServiceFile.info("listDicke $listDicke")
        logServiceFile.info("listRohdichte $listRohdichte")
        logServiceFile.info("listQuerzug $listQuerzug")
        logServiceFile.info("listAbhebe $listAbhebe")
        logServiceFile.info("listBiege $listBiege")
        logServiceFile.info("listEModul $listEModul")
        logServiceFile.info("listQuellung24h $listQuellung24h")
        logServiceFile.info("listRestfeuchte $listRestfeuchte")
        logServiceFile.info("listWasseraufnahme $listWasseraufnahme")



        readDataFromPLC(
            nameFilePLC,
            "Dryer" to listDryerPlc,
            "Refiner" to listRefinerPlc,
            "GlueKitchen" to listGlueKitchenPlc,
            "Dosing Chips" to listDosingChipsPlc,
            "Matformer & Press" to listMatformerPresssPlc,
            "Product Info" to listProductInfoPlc,
        )
        logServiceFile.info("Dryer $listDryerPlc")
        logServiceFile.info("Refiner $listRefinerPlc")
        logServiceFile.info("Dosing Chips $listDosingChipsPlc")
        logServiceFile.info("Matformer & Press  $listMatformerPresssPlc")
        logServiceFile.info("Product Info $listProductInfoPlc")

        TimeUnit.SECONDS.sleep(5)
        makeTableLabReport(
            listDicke, listRohdichte, listQuerzug, listAbhebe, listBiege,
            listEModul, listQuellung24h, listRestfeuchte, listWasseraufnahme,
            listProductInfoPlc, listRefinerPlc, listDosingChipsPlc, listGlueKitchenPlc,
            listDryerPlc, listMatformerPresssPlc
        )

    }

    //todo read data from excl plc document
    private fun readDataFromPLC(fileNamePlc: String?, vararg pairListPlc: Pair<String, MutableList<String>>) {

        val resourcePlc = javaClass.classLoader.getResource("upload_plc/$fileNamePlc")
        logServiceFile.info("resourcePlc.toURI() ${resourcePlc.toURI()}")

        val excelFilePlc = FileInputStream(File(resourcePlc.toURI()))
        val workBookPlc = XSSFWorkbook(excelFilePlc)
        val sheetPlc = workBookPlc.getSheet("Sheet1")

        pairListPlc.forEach {
            when (it.first) {
                "Dryer" -> it.second.addAll(
                    getDataFromPLC(sheetPlc, "Dryer", 'B', 8, 1)
                )
                "Refiner" -> it.second.addAll(
                    getDataFromPLC(sheetPlc, "Refiner", 'D', 7, 1)
                )
                "GlueKitchen" -> it.second.addAll(
                    getDataFromPLC(sheetPlc, "Glue Kitchen", 'F', 6, 1)
                )
                "Dosing Chips" -> it.second.addAll(
                    getDataFromPLC(sheetPlc, "Dosing Chips", 'H', 3, 1)
                )
                "Matformer & Press" -> it.second.addAll(
                    getDataFromPLC(sheetPlc, "Matformer & Press", 'J', 10, 1)
                )
                "Product Info" -> it.second.addAll(
                    getDataFromPLC(sheetPlc, "Product", 'L', 6, 1)
                )

            }
        }

    }

    private fun makeTableLabReport(vararg listDataValueFromIMAL_and_PLC: MutableList<String>) {
        println("call makeTableLabReport  ${listDataValueFromIMAL_and_PLC.size}")

        val workBook = XSSFWorkbook()
        val createHelper = workBook.creationHelper

        val sheet = workBook.createSheet("Customer")

        sheet.setColumnWidth(0, 1000)
        sheet.setColumnWidth(1, 2000)
        sheet.setColumnWidth(2, 4000)
        sheet.setColumnWidth(3, 4000)
        sheet.setColumnWidth(4, 2000)


        //todo header title
        val headerTitle = workBook.createCellStyle()
        headerTitle.setVerticalAlignment(VerticalAlignment.CENTER)
        headerTitle.setAlignment(HorizontalAlignment.CENTER)
        headerTitle.setFillForegroundColor(IndexedColors.AQUA.index)
        headerTitle.setFillPattern(FillPatternType.SOLID_FOREGROUND)
        val rowTitle = sheet.createRow(1)
        val cellTitle = rowTitle.createCell(1)
        cellTitle.setCellValue("Fiberboard")
        sheet.addMergedRegion(CellRangeAddress(0, 1, 1, 10))
        cellTitle.cellStyle = headerTitle
//todo font
        val fontForPLCTable = workBook.createFont()
        fontForPLCTable.setFontHeight(9.0)
        val fontTable = workBook.createFont()
        fontTable.setFontHeight(10.0)
        //todo general style
        val generalStyle = workBook.createCellStyle()
        generalStyle.setVerticalAlignment(VerticalAlignment.TOP)
        generalStyle.setAlignment(HorizontalAlignment.CENTER)
        generalStyle.setBorderTop(BorderStyle.THIN)
        generalStyle.setBorderBottom(BorderStyle.THIN)
        generalStyle.setBorderLeft(BorderStyle.THIN)
        generalStyle.setBorderRight(BorderStyle.THIN)
        generalStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.index)
        generalStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)


        val generalStyleTable = workBook.createCellStyle()
        generalStyleTable.setVerticalAlignment(VerticalAlignment.TOP)
        generalStyleTable.setAlignment(HorizontalAlignment.CENTER)
        generalStyleTable.setBorderTop(BorderStyle.THIN)
        generalStyleTable.setBorderBottom(BorderStyle.THIN)
        generalStyleTable.setBorderLeft(BorderStyle.THIN)
        generalStyleTable.setBorderRight(BorderStyle.THIN)
        generalStyleTable.setFillForegroundColor(IndexedColors.AQUA.index)
        generalStyleTable.setFillPattern(FillPatternType.SOLID_FOREGROUND)
        generalStyleTable.setFont(fontTable)

        val generalStyleFromPLC_Left = workBook.createCellStyle()
        generalStyleFromPLC_Left.setVerticalAlignment(VerticalAlignment.TOP)
        generalStyleFromPLC_Left.setAlignment(HorizontalAlignment.CENTER)
        generalStyleFromPLC_Left.setBorderLeft(BorderStyle.THIN)
        generalStyleFromPLC_Left.setFont(fontForPLCTable)
        generalStyleFromPLC_Left.setFillForegroundColor(IndexedColors.AQUA.index)
        generalStyleFromPLC_Left.setFillPattern(FillPatternType.SOLID_FOREGROUND)


        val generalStyleFromPLC_Right = workBook.createCellStyle()
        generalStyleFromPLC_Right.setVerticalAlignment(VerticalAlignment.TOP)
        generalStyleFromPLC_Right.setAlignment(HorizontalAlignment.CENTER)
        generalStyleFromPLC_Right.setBorderRight(BorderStyle.THIN)
        generalStyleFromPLC_Right.setFont(fontForPLCTable)
        generalStyleFromPLC_Right.setFillForegroundColor(IndexedColors.SEA_GREEN.index)
        generalStyleFromPLC_Right.setFillPattern(FillPatternType.SOLID_FOREGROUND)


        //todo fusion cell -----------------------------------------------------------------
        sheet.addMergedRegion(CellRangeAddress(2, 3, 1, 2))
// todo init header one
        val listHeaderTextOne = listOf<String>("Charge: \n", "Typ: \n", "Prod.zeit: \n", "Pruf.Nr.: \n")
        val listHeaderTextValuesOne = listOf<String>("123456", "HDF", LocalDateTime.now().toString(), "79563466740")
        //todo text and value 1------------------------------------------------------------------------
        val rowHeaderOne = sheet.createRow(2)

        createHeaderOne(sheet, rowHeaderOne, 0, 1, 1, generalStyle, listHeaderTextOne, listHeaderTextValuesOne)
        //todo text and value 2------------------------------------------------------------------------
        createHeaderOne(sheet, rowHeaderOne, 1, 3, 4, generalStyle, listHeaderTextOne, listHeaderTextValuesOne)
        // todo text and value 3------------------------------------------------------------------------
        createHeaderOne(sheet, rowHeaderOne, 2, 5, 7, generalStyle, listHeaderTextOne, listHeaderTextValuesOne)
        //  todo text and value 4------------------------------------------------------------------------
        createHeaderOne(sheet, rowHeaderOne, 3, 8, 10, generalStyle, listHeaderTextOne, listHeaderTextValuesOne)

        rowHeaderOne.createCell(12).setCellValue("Svetikov")

//todo -----------------------------------------------------------------------------------------------------------------------------------------------------------------------
//todo --Nr--|--Dicke--|--Rohdichte--|--Querzugfestigkeit--|--Abhebe-festigkell--|--Beige-festigkell--|--E-Modul-quer--|--Quellung-24h--|--Rest-feuchte--|Wasser-aufnahme----|--------------------------------------
//todo -----------------------------------------------------------------------------------------------------------------------------------------------------------------------
        val listHeaderTwoText = listOf<String>(
            " Nr ",
            "Dicke",
            "Rohdichte",
            "Querzug-\n festigkeit",
            "Abhebe-\n festigkell",
            "Beige-\n festigkell",
            "E-Modul-\n quer",
            "Quellung- \n 24h",
            "Rest- \n feuchte",
            "Wasser- \n aufnahme"
        )

        val listHeaderTwoSI = listOf<String>(
            "_____", "mm", "kg/m2", "N/mm2", "N/mm2", "N/mm2", "N/mm2", "%", "%", "%"
        )
        val listNumberTwo = listOf<String>("1", "2", "3", "4", "5", "6", "7", "8", "MW", "SOLL", "Max", "Min")


        //todo header text
        val rowHeaderTwoText = sheet.createRow(5)
        var cellHeaderTwoText = rowHeaderTwoText.createCell(1)
        //todo header si
        val rowHeaderTwoSI = sheet.createRow(7)
        var cellHeaderTwoSI = rowHeaderTwoSI.createCell(1)
        //todo table for value from laboratory
        var rowValueFromLaboratory = sheet.createRow(18)
        var cellValueFromLaboratory = rowValueFromLaboratory.createCell(1)

        for (index in listNumberTwo.indices) { //listHeaderTwoText.indices
            //todo create header text

            if (index < 10) {
                cellHeaderTwoText = rowHeaderTwoText.createCell(index + 1)
                cellHeaderTwoText.setCellValue(listHeaderTwoText[index])
                cellHeaderTwoText.cellStyle = generalStyle
            }
            //todo create header SI

            if (index < 10) {
                cellHeaderTwoSI = rowHeaderTwoSI.createCell(index + 1)
                cellHeaderTwoSI.setCellValue(listHeaderTwoSI[index])
                cellHeaderTwoSI.cellStyle = generalStyle

            }

            sheet.autoSizeColumn(index + 1)
            if (index < 10) {
                sheet.addMergedRegion(
                    CellRangeAddress(
                        5,
                        6,
                        cellHeaderTwoText.columnIndex,
                        cellHeaderTwoText.columnIndex
                    )
                )
            }
            //todo create header value laboratory
            rowValueFromLaboratory = sheet.createRow(index + 8)
            for (i in listHeaderTwoSI.indices) {
//todo create and fulling cells data from imal laboratory
                cellValueFromLaboratory = rowValueFromLaboratory.createCell(i + 1)
                when (i) {
                    0 -> cellValueFromLaboratory.setCellValue(listNumberTwo[index])
                    1 -> cellValueFromLaboratory.setCellValue(listDataValueFromIMAL_and_PLC[0][index])
                    2 -> cellValueFromLaboratory.setCellValue(listDataValueFromIMAL_and_PLC[1][index])
                    3 -> cellValueFromLaboratory.setCellValue(listDataValueFromIMAL_and_PLC[2][index])
                    4 -> cellValueFromLaboratory.setCellValue(listDataValueFromIMAL_and_PLC[3][index])
                    5 -> cellValueFromLaboratory.setCellValue(listDataValueFromIMAL_and_PLC[4][index])
                    6 -> cellValueFromLaboratory.setCellValue(listDataValueFromIMAL_and_PLC[5][index])
                    7 -> cellValueFromLaboratory.setCellValue(listDataValueFromIMAL_and_PLC[6][index])
                    8 -> cellValueFromLaboratory.setCellValue(listDataValueFromIMAL_and_PLC[7][index])
                    9 -> cellValueFromLaboratory.setCellValue(listDataValueFromIMAL_and_PLC[8][index])
                    else -> throw NoSuchElementException("No list with iteration $i")
                }
                cellValueFromLaboratory.cellStyle = generalStyleTable
            }
        }


//todo middle general data (allgemeine)
        val listHeaderData = listOf<String>(
            "Product Data",
            "Refiner",
            "Dosing Chips", "Glue Kitchen", "Dryer", "Matformer & Press"
        )

        var step = 21 //todo start row
        var bias = 1
        var bias2 = 3
        var rowHeaderDataPLC = sheet.createRow(step)
        for (index in listHeaderData.indices) {
            if (index == 3) { // todo jump to next row
                step += 15
                bias = 1
                bias2 = 3
                rowHeaderDataPLC = sheet.createRow(step)
            }
            val cellHeaderDataPLC = rowHeaderDataPLC.createCell(bias)
            cellHeaderDataPLC.setCellValue(listHeaderData[index])

            if (index == 0 || index == 3) { //todo rice 4 cell first note
                bias2 += 1
                //println("1 $bias // $bias2 //$index")
                sheet.addMergedRegion(CellRangeAddress(step, step, bias, bias2))
                bias = bias2 + 1
                bias2 += 3
                // println("2 $bias // $bias2 //$index")
            }

            if (index != 0 && index != 3) { //todo other cell note
                //   println("3 $bias // $bias2 //$index")
                sheet.addMergedRegion(CellRangeAddress(step, step, bias, bias2))
                bias = bias2 + 1
                bias2 += 3
                //  println("4 $bias // $bias2 //$index")
            }
            cellHeaderDataPLC.cellStyle = generalStyle
        }
        //todo value from PLC

        val listTextProductInfoPLC =
            listOf(
                "", "Recipe name", "Product code", "Thickness", "Width", "Length", "Density",
                "", "", "", "", "", "", ""
            )
        val listTextRefinerInfoPLC =
            listOf(
                "",
                "Preheater temp",
                "Preheater stream",
                "Difibrator Pressure",
                "Prehiater hoist min",
                "Motor power SEC",
                "Production t/h",
                "Blow valve %",
                "",
                "",
                "",
                "",
                "",
                ""
            )
        val listTextDosingChipsPLC =
            listOf("", "Silos 1", "Silos 2", "Imported chips", "", "", "", "", "", "", "", "", "", "")
        val listTextGlueKitchenPLC =
            listOf("", "UF Glue", "MUF Glue", "Water", "Hardener", "Urea", "Emulsion", "", "", "", "", "", "", "")
        val listTextDryerPLC =
            listOf(
                "", "Temp inlet", "Pres inlet", "Temp pipe", "Moisture", "Temp cyclone 1", "Temp cyclone 2",
                "Filling level bunker", "Dosing bunker", "", "", "", "", ""
            )
        val listTextMatformerPressPLC =
            listOf(
                "",
                "Temp bunker",
                "Density calculation",
                "Weight on scale",
                "Heating plate 1",
                "Heating plate 2",
                "Heating plate 3",
                "Heating plate 4",
                "Heating plate 5",
                "Speed press",
                "Press factor",
                "",
                "",
                ""
            )

        val listValueProductPLC = listOf(
            "",
            listDataValueFromIMAL_and_PLC[9][0],
            listDataValueFromIMAL_and_PLC[9][1],
            listDataValueFromIMAL_and_PLC[9][2] + "mm",
            listDataValueFromIMAL_and_PLC[9][3] + "mm",
            listDataValueFromIMAL_and_PLC[9][4] + "mm",
            listDataValueFromIMAL_and_PLC[9][5] + "kg/m3",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
        )
        val listValueRefinerPLC = listOf(
            "",
            listDataValueFromIMAL_and_PLC[10][0],
            listDataValueFromIMAL_and_PLC[10][1],
            listDataValueFromIMAL_and_PLC[10][2],
            listDataValueFromIMAL_and_PLC[10][3],
            listDataValueFromIMAL_and_PLC[10][4],
            listDataValueFromIMAL_and_PLC[10][5],
            listDataValueFromIMAL_and_PLC[10][6],
            "",
            "",
            "",
            "",
            "",
            ""
        )
        val listValueDosingChipsPLC = listOf(
            "",
            listDataValueFromIMAL_and_PLC[11][0],
            listDataValueFromIMAL_and_PLC[11][1],
            listDataValueFromIMAL_and_PLC[11][2],
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
        )
        val listValueGlueKitchenPLC = listOf(
            "",
            listDataValueFromIMAL_and_PLC[12][0],
            listDataValueFromIMAL_and_PLC[12][1],
            listDataValueFromIMAL_and_PLC[12][2],
            listDataValueFromIMAL_and_PLC[12][3],
            listDataValueFromIMAL_and_PLC[12][4],
            listDataValueFromIMAL_and_PLC[12][5],
            "",
            "",
            "",
            "",
            "",
            "",
            ""
        )
        val listValueDryerPLC = listOf(
            "",
            listDataValueFromIMAL_and_PLC[13][0],
            listDataValueFromIMAL_and_PLC[13][1],
            listDataValueFromIMAL_and_PLC[13][2],
            listDataValueFromIMAL_and_PLC[13][3],
            listDataValueFromIMAL_and_PLC[13][4],
            listDataValueFromIMAL_and_PLC[13][5],
            listDataValueFromIMAL_and_PLC[13][6],
            listDataValueFromIMAL_and_PLC[13][7],
            "",
            "",
            "",
            "",
            ""
        )
        val listValueMatformerPress = listOf(
            "",
            listDataValueFromIMAL_and_PLC[14][0],
            listDataValueFromIMAL_and_PLC[14][1],
            listDataValueFromIMAL_and_PLC[14][2],
            listDataValueFromIMAL_and_PLC[14][3],
            listDataValueFromIMAL_and_PLC[14][4],
            listDataValueFromIMAL_and_PLC[14][5],
            listDataValueFromIMAL_and_PLC[14][6],
            listDataValueFromIMAL_and_PLC[14][7],
            listDataValueFromIMAL_and_PLC[14][8],
            listDataValueFromIMAL_and_PLC[14][9],
            "",
            "",
            ""
        )
        var biasPLC = 22
        for (index in listValueProductPLC.indices) {

            var rowValueFromPLC = sheet.createRow(biasPLC)
            cellFull(
                sheet, biasPLC, rowValueFromPLC, index, 1, 2, 3, 4,
                generalStyleFromPLC_Left, generalStyleFromPLC_Right, listTextProductInfoPLC, listValueProductPLC
            )
            cellFull(
                sheet, biasPLC, rowValueFromPLC, index, 5, 6, 7, 7,
                generalStyleFromPLC_Left, generalStyleFromPLC_Right, listTextRefinerInfoPLC, listValueRefinerPLC
            )
            cellFull(
                sheet, biasPLC, rowValueFromPLC, index, 8, 9, 10, 10,
                generalStyleFromPLC_Left, generalStyleFromPLC_Right, listTextDosingChipsPLC, listValueDosingChipsPLC
            )

            biasPLC += 1

        }
        biasPLC = 37
        for (index in listValueProductPLC.indices) {

            var rowValueFromPLC = sheet.createRow(biasPLC)
            cellFull(
                sheet, biasPLC, rowValueFromPLC, index, 1, 2, 3, 4,
                generalStyleFromPLC_Left, generalStyleFromPLC_Right, listTextGlueKitchenPLC, listValueGlueKitchenPLC
            )
            cellFull(
                sheet, biasPLC, rowValueFromPLC, index, 5, 6, 7, 7,
                generalStyleFromPLC_Left, generalStyleFromPLC_Right, listTextDryerPLC, listValueDryerPLC
            )
            cellFull(
                sheet, biasPLC, rowValueFromPLC, index, 8, 9, 10, 10,
                generalStyleFromPLC_Left, generalStyleFromPLC_Right, listTextMatformerPressPLC, listValueMatformerPress
            )

            biasPLC += 1

        }
        var biasBottom = 51
        for (index in 1..5) {
            val rowBottom = sheet.createRow(biasBottom)
            val cellBottom = rowBottom.createCell(1)
            biasBottom += 1
            cellBottom.cellStyle = generalStyle
        }

        sheet.addMergedRegion(CellRangeAddress(51, biasBottom, 1, 10))

        val formatter = DateTimeFormatter.ofPattern("d-MM-yyyy-HH-mm-ss")
        val date = LocalDateTime.now().format(formatter)
        path = "D:\\customer ${date}.xlsx"

        val fileOut = FileOutputStream(path)
        workBook.write(fileOut)
        fileOut.close()
        workBook.close()

        Runtime.getRuntime().exec("explorer $path")


        listDataValueFromIMAL_and_PLC.forEach { it.clear() }
    }

    private fun readTableFromIMAL(fileName: String?, vararg pairList: Pair<String, MutableList<String>>) {
        logServiceFile.info("readTableFromIMAL(fileName $fileName: String?, vararg pairList $pairList: Pair<String, MutableList<String>>)")
        val resources = javaClass.classLoader.getResource("upload/$fileName")
        logServiceFile.info("resources.toURI() ${resources.toURI()}")
        val excelFile = FileInputStream(File(resources.toURI()))//my-server/src/main/resources/upload
        val workBook = HSSFWorkbook(excelFile)
        val sheet = workBook.getSheet("IMAL Test")

        pairList.forEach {
            when (it.first) {
                "Dicke" -> it.second.addAll(
                    addAverageMaxMin(
                        getDataFromIMAL(
                            sheet,
                            "Растяжение",
                            'D',
                            8,
                            4
                        )
                    )
                )
                "Rohdichte" -> it.second.addAll(
                    addAverageMaxMin(
                        getDataFromIMAL(
                            sheet,
                            "Растяжение",
                            'F',
                            8,
                            4
                        )
                    )
                )
                "Querzug" -> it.second.addAll(
                    addAverageMaxMin(
                        getDataFromIMAL(
                            sheet,
                            "Растяжение",
                            'I',
                            8,
                            4
                        )
                    )
                )
                "Abhebe" -> it.second.addAll(
                    addAverageMaxMinSkip(
                        2,
                        getDataFromIMAL(sheet, "Поверхн.растяжение", 'I', 6, 4)
                    )
                )
                "Biege" -> it.second.addAll(
                    addAverageMaxMinSkip(
                        2,
                        getDataFromIMAL(sheet, "Изгиб", 'I', 6, 2)
                    )
                )
                "EModul" -> it.second.addAll(
                    addAverageMaxMinSkip(
                        2,
                        getDataFromIMAL(sheet, "Мод.Упр.", 'J', 6, 2)
                    )
                )
                "Quellung24h" -> it.second.addAll(
                    addAverageMaxMinSkip(
                        2,
                        getDataFromIMAL(sheet, "Разбухание", 'K', 6, 3)
                    )
                )
                "Restfeuchte" -> it.second.addAll(
                    addAverageMaxMinSkip(
                        4,
                        getDataFromIMAL(sheet, "Влажность", 'I', 4, 2)
                    )
                )
                "Wasseraufnahme" -> it.second.addAll(
                    addAverageMaxMinSkip(
                        2,
                        getDataFromIMAL(sheet, "Влагопоглощение", 'K', 6, 3)
                    )
                )
                else -> throw NoSuchElementException(it.first)
            }
        }
    }

    fun getDataFromIMAL(sheet: HSSFSheet, name: String, cell: Char, getRows: Int, offset: Int): MutableList<Double> {
        val cellNumber = ('A'..'Z').toList().indexOf(cell)
        var list: List<String> = sheet.asSequence().map { it.getCell(cellNumber) }
            .filterNotNull()
            .map { it.toString() }
            .filter { it.isNotEmpty() }
            .toList()
        val index = list.indexOfLast { it == name } + offset
        if (index - offset != -1)
            list = list.subList(index, index + getRows)
        else {
            list = (1..8).toList() as MutableList<String>
            list.fill("1")
        }

        return list.map { it.toDouble() } as MutableList<Double>
    }

    fun getDataFromPLC(sheet: XSSFSheet, name: String, cell: Char, getRows: Int, offset: Int): MutableList<String> {
        val cellNumber = ('A'..'Z').toList().indexOf(cell)
        var list: List<String> = sheet.asSequence().map { it.getCell(cellNumber) }
            .filterNotNull()
            .map { it.toString() }
            .filter { it.isNotEmpty() }
            .toList()
        val index = list.indexOfLast { it == name } + offset
        if (index - offset != -1)
            list = list.subList(index, index + getRows)
        else {
            list = (1..11).toList() as MutableList<String>
            list.fill("1")
        }

        return list as MutableList<String>
    }


    private fun addAverageMaxMin(list: MutableList<Double>): MutableList<String> {
        val aver = (list.average() * 100).roundToInt() / 100.0
        val std = list.map { (it - aver).pow(2.0) }.reduce { acc, it -> it + acc }.div(list.size.toDouble())
        val max = list.maxOrNull()
        val min = list.minOrNull()
        list.add(aver)
        list.add((sqrt(std)))
        list.add(max!!)
        list.add(min!!)
        return list.map { (it * 100).roundToInt() / 100.0 }.map { it.toString() } as MutableList<String>
    }

    private fun addAverageMaxMinSkip(skip: Int, list: MutableList<Double>): MutableList<String> {
        val aver = (list.average() * 100).roundToInt() / 100.0
        val std = list.map { (it - aver).pow(2.0) }.reduce { acc, it -> it + acc }.div(list.size.toDouble())
        val max = list.maxOrNull()
        val min = list.minOrNull()

        val listLocal = (1..skip).toMutableList()
        listLocal.fill(0)
        list.addAll(listLocal.map { it.toDouble() })

        list.add(aver)
        list.add((sqrt(std)))
        list.add(max!!)
        list.add(min!!)
        return list.map { (it * 100).roundToInt() / 100.0 }.map { it.toString() } as MutableList<String>

    }

    fun getListPLC(sheet: XSSFSheet, name: String, cell: Char, getRows: Int, offset: Int): List<String> {
        val cellNumber = ('A'..'Z').toList().indexOf(cell)
        var list: List<String> = sheet.asSequence().map { it.getCell(cellNumber) }
            .filterNotNull()
            .map { it.toString() }
            .filter { it.isNotEmpty() }
            .toList()
        val index = list.indexOfLast { it == name } + offset
        list = list.subList(index, index + getRows)
        return list
    }

    private fun cellFull(
        sheet: XSSFSheet,
        rowIndex: Int,
        rowStart: XSSFRow,
        index: Int,
        cellOne: Int,
        biasOne: Int,
        cellTwo: Int,
        biasTwo: Int,
        styleOne: XSSFCellStyle,
        styleTwo: XSSFCellStyle,
        listText: List<String>,
        listValue: List<String>
    ) {
        val cellText = rowStart.createCell(cellOne)
        val cellValue = rowStart.createCell(cellTwo)
        cellText.setCellValue(listText[index])
        cellValue.setCellValue(listValue[index])

        cellText.cellStyle = styleOne
        cellValue.cellStyle = styleTwo
        if (cellOne != biasOne)
            sheet.addMergedRegion(CellRangeAddress(rowIndex, rowIndex, cellOne, biasOne))
        if (cellTwo != biasTwo)
            sheet.addMergedRegion(CellRangeAddress(rowIndex, rowIndex, cellTwo, biasTwo))
    }


    //todo createHeaderOne
    private fun createHeaderOne(
        sheet: XSSFSheet,
        row: XSSFRow,
        index: Int,
        cellOne: Int,
        cellTwo: Int,
        generalStyle: XSSFCellStyle,
        listText: List<String>,
        listValue: List<String>
    ) {
        val cellHeaderOne1 = row.createCell(cellOne)

        if (cellOne != cellTwo)
            sheet.addMergedRegion(CellRangeAddress(2, 3, cellOne, cellTwo))
        cellHeaderOne1.setCellValue(listText[index] + listValue[index])
        cellHeaderOne1.cellStyle = generalStyle
    }


}