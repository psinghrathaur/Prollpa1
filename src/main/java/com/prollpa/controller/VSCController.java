package com.prollpa.controller;

import java.util.List;

import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import com.prollpa.entity.VSC;
import com.prollpa.service.VSCService;

import io.swagger.v3.oas.annotations.tags.Tag;

@RestController
@RequestMapping("/vsc")
@Tag(name = "VSC Controller API", description = "VSC Management")
public class VSCController {
	private final VSCService vscService;
	public VSCController(VSCService vscService) {
		this.vscService=vscService;
	}
	@PostMapping("save")
	public ResponseEntity<VSC> saveVSC(@RequestBody VSC vsc){
		VSC saveVSC = vscService.saveVSC(vsc);
		return ResponseEntity.status(HttpStatus.CREATED).body(saveVSC);
	}
	@GetMapping("/getVSCById")
	public ResponseEntity<VSC> getVSCById(@RequestParam("vscId") long vscId){
		VSC vsc = vscService.getVSCById(vscId);
		return ResponseEntity.status(HttpStatus.OK).body(vsc);
	}
	@GetMapping("/getVSCList")
	public ResponseEntity<List<VSC>> getVSCList(){
		List<VSC> vscList = vscService.getVSCList();
		return ResponseEntity.status(HttpStatus.OK).body(vscList);
	}
	

}
