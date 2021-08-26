package com.svetikov.myserver.controller

import org.springframework.web.bind.annotation.CrossOrigin
import org.springframework.web.bind.annotation.GetMapping
import org.springframework.web.bind.annotation.RequestMapping
import org.springframework.web.bind.annotation.RestController
@CrossOrigin("http://localhost:4200")
@RestController
@RequestMapping("/api/lab")
class MyController {
    @GetMapping
    fun test() = Message("Hello")
}

data class Message(val mess:String)