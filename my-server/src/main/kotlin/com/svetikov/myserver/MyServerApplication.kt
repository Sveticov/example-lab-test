package com.svetikov.myserver

import org.springframework.boot.autoconfigure.SpringBootApplication
import org.springframework.boot.runApplication
import org.springframework.boot.web.servlet.MultipartConfigFactory
import org.springframework.context.annotation.Bean
import org.springframework.util.unit.DataSize
import org.springframework.web.bind.annotation.RestController
import javax.servlet.MultipartConfigElement

@SpringBootApplication

class MyServerApplication

fun main(args: Array<String>) {
    runApplication<MyServerApplication>(*args)
}

@Bean
fun multipartConfigureElement():MultipartConfigElement{
    val factory = MultipartConfigFactory()
    factory.setMaxFileSize(DataSize.ofKilobytes(512))
    factory.setMaxRequestSize(DataSize.ofKilobytes(512))
    return factory.createMultipartConfig()
}
