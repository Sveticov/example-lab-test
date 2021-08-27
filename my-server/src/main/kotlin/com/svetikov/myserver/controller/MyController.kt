package com.svetikov.myserver.controller

import com.svetikov.myserver.service.ServiceFile
import org.apache.catalina.loader.ResourceEntry
import org.springframework.http.HttpStatus
import org.springframework.http.ResponseEntity
import org.springframework.web.bind.annotation.*
import org.springframework.web.multipart.MultipartFile

@CrossOrigin("http://localhost:4200","/files")
@RestController
@RequestMapping("/api/lab")
class MyController(private val service:ServiceFile) {
    val files = mutableListOf<File>(File("f1",12))
    @GetMapping
    fun test() = Message("Hello")

    @PostMapping("/filet")
    fun getFile(@RequestBody file:File):Boolean{

        val fileExist = file.exist()
        println("$file and exist $fileExist")
        if (fileExist)
        files.add(file)
        return fileExist
    }

    @GetMapping("/list")
    fun getAllList() = files


    @PostMapping("/files")
    fun uploadFile(@RequestParam("file") files:MultipartFile):ResponseEntity<File>{
        println(files.originalFilename)
        service.save(files)
        return ResponseEntity.status(HttpStatus.OK).body(File(files.originalFilename!!,files.size))
    }

}

data class Message(val mess:String)
data class File(val name:String,val size:Long) {
    fun exist(): Boolean {
if (!name.isNullOrBlank() && size!=null) return true
        return false
    }
}