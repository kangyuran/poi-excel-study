package com.example.poiexcelstudy.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.ResponseBody;

@Controller
public class TestController {

    @ResponseBody
    public String test() {
        return "OK";
    }
}
