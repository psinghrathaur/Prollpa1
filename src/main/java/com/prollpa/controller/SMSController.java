package com.prollpa.controller;

import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import com.prollpa.externalApis.SMS.SMSService;
import com.prollpa.payload.SMSRequest;

@RestController
@RequestMapping("/sms")
public class SMSController {
	private final SMSService smsService;

    public SMSController(SMSService smsService) {
        this.smsService = smsService;
    }

    @PostMapping("/send-sms")
    public String sendSms(@RequestBody SMSRequest smsDto ) {
    	System.out.println(smsDto.getMessage()+"->"+smsDto.getMobileNumber());
        return smsService.sendSms(smsDto.getMobileNumber(), smsDto.getMessage());
    }
    @GetMapping("/getSMSStatusById/{smsId}")
    public ResponseEntity<String> getSMSStatusById(@PathVariable("smsId")String smsId){
    	
    	return smsService.getSmsStatus(smsId); 
    	
    }

}
