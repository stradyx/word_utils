package com.strady.word_utils.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

/**
 * @Author: strady
 * @Date: 2020-01-04
 * @Time: 20:36:44
 * @Description:
 */
@Controller
@RequestMapping(value = "/")
public class ViewController {

    @RequestMapping(value = "/index")
    public String index() {
        return "index";
    }
}
