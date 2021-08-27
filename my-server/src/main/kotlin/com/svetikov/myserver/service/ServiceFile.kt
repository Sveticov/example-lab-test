package com.svetikov.myserver.service

import org.apache.tomcat.util.http.fileupload.FileUploadException
import org.springframework.beans.factory.annotation.Value
import org.springframework.stereotype.Service
import org.springframework.web.multipart.MultipartFile
import java.lang.Exception
import java.nio.file.Files
import java.nio.file.Paths
import kotlin.io.path.deleteIfExists

@Service
class ServiceFile {

    @Value("\${upload.path}")
    lateinit var uploadPath:String//my-server/src/main/resources/upload

    fun save(file: MultipartFile) {
        try {
            val root = Paths.get(uploadPath)
            val resolve = root.resolve(file.originalFilename)
            if (resolve.toFile().exists()) {
              val status =  resolve.deleteIfExists()
               if (!status)
                throw FileUploadException("File already exists ${file.originalFilename}")
            }
            Files.copy(file.inputStream, resolve)
        } catch (e: Exception) {
            throw FileUploadException("Couldn't store the file.Error ${e.message}")
        }
    }



}