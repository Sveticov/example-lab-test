package com.svetikov.myserver.controller

import com.svetikov.myserver.service.ServiceFile
import org.apache.catalina.loader.ResourceEntry
import org.springframework.http.HttpStatus
import org.springframework.http.ResponseEntity
import org.springframework.web.bind.annotation.*
import org.springframework.web.multipart.MultipartFile

@CrossOrigin("http://localhost:4200", "/files")
@RestController
@RequestMapping("/api/lab")
class MyController(private val service: ServiceFile) {
    val files = mutableListOf<File>(File("f1", 12))

    @GetMapping
    fun test() = Message("Hello")

    @PostMapping("/filet")
    fun getFile(@RequestBody file: File): Boolean {
        val fileExist = file.exist()
        println("$file and exist $fileExist")
        if (fileExist)
            files.add(file)
        return fileExist
    }

    @GetMapping("/list")
    fun getAllList() = files

    //todo upload imal excel data
    @PostMapping("/files")
    fun uploadFile(@RequestParam("file") files: MultipartFile): ResponseEntity<File> {
        service.save(files)
        return ResponseEntity.status(HttpStatus.OK).body(File(files.originalFilename!!, files.size))
    }

    //todo upload plc simatic excel data
    @PostMapping("/files_plc")
    fun uploadFilePLC(@RequestParam("file_plc") files_plc: MultipartFile): ResponseEntity<File> {
        println(files_plc.originalFilename)
        service.savePlc(files_plc)
        return ResponseEntity.status(HttpStatus.OK).body(File(files_plc.originalFilename!!, files_plc.size))
    }
    //todo code for working table EXCEL

    @GetMapping("/table")
    fun makeTableLaboratoryReport(): Boolean {
        println("call /table")
        service.readIMALAndCreateTableLaboratoryReport()
        return true
    }

}

data class Message(val mess: String)
data class File(val name: String, val size: Long) {
    fun exist(): Boolean {
        if (!name.isNullOrBlank() && size != null) return true
        return false
    }
}