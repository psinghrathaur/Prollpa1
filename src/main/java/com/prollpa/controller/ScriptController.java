package com.prollpa.controller;

import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.prollpa.serviceImpl.ScriptServiceImpl;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.core.io.Resource;
import org.springframework.core.io.Resource;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/Scripts")
public class ScriptController { 
	private ScriptServiceImpl scriptService;
	public ScriptController(ScriptServiceImpl scriptService) {
		
		this.scriptService = scriptService;
	}
	@GetMapping("/download-script")
    public ResponseEntity<Resource> downloadScript()  {
    try {
		return scriptService.generateScriptFile();
	} catch (InvalidFormatException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
    return null;
  }
  @PostMapping("/getnationalitytranslationScript")
  public ResponseEntity<Resource> getnationalitytranslationScript(){
	  try {
			return scriptService.getnationalitytranslationScript();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	  return null;
  }
  @PostMapping("/getLocalLanguageTranslationScript")
  public ResponseEntity<Resource> getLocalLanguageTranslationScript(){
	  try {
			return scriptService.getLocalLanguageTranslationScript();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	  return null;
  }
  

}
