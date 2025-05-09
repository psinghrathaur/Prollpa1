package com.prollpa.controller;

import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
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
@RequestMapping("api/scripts")
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
  
  @GetMapping("/generatecountryvisalinkUATOutput")
  public ResponseEntity<Resource> generatecountryvisalinkUATOutput()  {
  try {
		return scriptService.generatecountryvisalinkUATOutput();
	} catch (InvalidFormatException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
  return null;
}


  @GetMapping("/generateNewLangNationalityUATOutput")
  public ResponseEntity<Resource> generateNewLangNationalityUATOutput()  {
  try {
		return scriptService.generateNewLangNationalityUATOutput();
	} catch (InvalidFormatException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
  return null;
}

  
  @GetMapping("/generateNewLangCommonUATLiveOutput")
  public ResponseEntity<Resource> generateNewLangCommonUATLiveOutput()  {
  try {
		return scriptService.generateNewLangNationalityUATOutput();
	} catch (InvalidFormatException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
  return null;
}

  @GetMapping("/generateVisaCountryDocLinkOutputUAT")
  public ResponseEntity<Resource> generateVisaCountryDocLinkOutputUAT()  {
  try {
		return scriptService.generateNewLangNationalityUATOutput();
	} catch (InvalidFormatException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
  return null;
}
  
  @GetMapping("/generatecountryvisalinkLiveOutput")
  public ResponseEntity<Resource> generatecountryvisalinkLiveOutput()  {
  try {
		return scriptService.generatecountryvisalinkLiveOutput();
	} catch (InvalidFormatException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
  return null;
}
  @GetMapping("/generateVSCVisaInsuranceUATOutput")
  public ResponseEntity<Resource> generateVSCVisaInsuranceUATOutput()  {
  try {
		return scriptService.generatecountryvisalinkLiveOutput();
	} catch (InvalidFormatException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
  return null;
}
  
  
  @GetMapping("/generateVSCVisaInsuranceLiveOutput")
  public ResponseEntity<Resource> generateVSCVisaInsuranceLiveOutput()  {
  try {
		return scriptService.generatecountryvisalinkLiveOutput();
	} catch (InvalidFormatException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
  return null;
}
 
  @GetMapping("/generateCourierCityOutput")
  public ResponseEntity<Resource> generateCourierCityOutput()  {
  try {
		return scriptService.generatecountryvisalinkLiveOutput();
	} catch (InvalidFormatException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
  return null;
}

  @GetMapping("/generateSMSforLocalLanguageOutput")
  public ResponseEntity<Resource> generateSMSforLocalLanguageOutput()  {
  try {
		return scriptService.generatecountryvisalinkLiveOutput();
	} catch (InvalidFormatException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
  return null;
}
  
  @GetMapping("/generateHolidayScriptOutput")
  public ResponseEntity<Resource> generateHolidayScriptOutput()  {
  try {
		return scriptService.generatecountryvisalinkLiveOutput();
	} catch (InvalidFormatException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
  return null;
}
  @PostMapping("/getAllScript")
  public ResponseEntity<String> getAllScript( @RequestParam("jsonRequiredList") String jsonRequiredList1) {
      System.out.println("Received: " + jsonRequiredList1);  // Check in console
      return ResponseEntity.ok("Server got: " + jsonRequiredList1);
  }
}
