package com.prollpa.controller;

import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.ResponseEntity;
import org.springframework.validation.BindingResult;
import org.springframework.validation.annotation.Validated;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.prollpa.entity.UserVscRole;
import com.prollpa.exception.ResourceNotFoundException;
import com.prollpa.payload.UserVscRoleDto;
import com.prollpa.payload.UserVscRoleResponse;
import com.prollpa.service.UserVscRoleService;
import com.prollpa.serviceImpl.UserVscRoleServiceImpl;

import jakarta.validation.Valid;

@RestController
@Validated
@RequestMapping("/userVscRole")
public class UserVscRoleController {
	private static final Logger logger = LoggerFactory.getLogger(UserVscRoleController.class);
	private final UserVscRoleService userVscRoleService;
	public UserVscRoleController(UserVscRoleService userVscRoleService) {
		this.userVscRoleService=userVscRoleService;
	}
	@PostMapping("/save")
    public ResponseEntity<?> saveUserVscRole(
            @RequestBody @Valid List<UserVscRoleDto> userVscRoleDtoList, BindingResult bindingResult) {
		   logger.info("Received UserVscRoleDto list: {}", userVscRoleDtoList);
		   
		    // If there are validation errors
		    if (bindingResult.hasErrors()) {
		        // Collect all error messages
		        String errorMessages = bindingResult.getAllErrors().stream()
		                .map(error -> error.getDefaultMessage())
		                .reduce((msg1, msg2) -> msg1 + "; " + msg2)
		                .orElse("Validation failed");

		        // Return a bad request with validation errors
		        throw new ResourceNotFoundException(errorMessages);
		    }
		   

        // Save the roles if valid
        List<UserVscRoleDto> savedUserVscRoleDtoList = userVscRoleService.saveUserVscRoleList(userVscRoleDtoList);
        return ResponseEntity.ok(savedUserVscRoleDtoList);
    }
	
	@GetMapping("/getUserVscRoleByUserId/{userId}")
	public ResponseEntity<List<UserVscRoleDto>> getUserVscRoleByUserId(@PathVariable("userId")long userId){
		return ResponseEntity.ok(userVscRoleService.getUserVscRoleByUserId(userId));
	}
	@GetMapping("/getUserVscRoleByUserId1/{userId}")
	public ResponseEntity<List<UserVscRoleResponse>> getUserVscRoleByUserId1(@PathVariable("userId")long userId){
		List<UserVscRoleDto> userVscRoleByUserId = userVscRoleService.getUserVscRoleByUserId(userId);
		List<UserVscRoleResponse> userVscRoleByUserId2 = userVscRoleService.getUserVscRoleByUserId2(userVscRoleByUserId);
		return ResponseEntity.ok(userVscRoleByUserId2);
	}

}
