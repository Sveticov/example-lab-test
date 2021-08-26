package com.svetikov.myserver

import org.springframework.boot.autoconfigure.SpringBootApplication
import org.springframework.boot.runApplication
import org.springframework.web.bind.annotation.RestController

@SpringBootApplication

class MyServerApplication

fun main(args: Array<String>) {
    runApplication<MyServerApplication>(*args)
}


