package com.prollpa.controller;

import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import com.prollpa.helperclass.UserRoleResponse;
import com.prollpa.payload.UserDto;
import com.prollpa.service.UserService;

import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;

@RestController
@RequestMapping("/users")
@Tag(name = "User Controller Api", description = "Users Management")
public class UserController {
	
	private UserService userService;
	
	public UserController(UserService userService) {
		super();
		this.userService = userService;
	}

	@PostMapping("/findUserRoleByUserNameAndVSCID")
	@Operation(summary = "findUserRoleByUserNameAndVSCID", description = "This Api will find the User Role for Perticular VSC")
	public ResponseEntity<UserRoleResponse> findUserRoleByUserNameAndVSCID(@PathVariable("userName")String userName){
		UserRoleResponse userRoleResponse=new UserRoleResponse();
		return ResponseEntity.ok(userRoleResponse);
	}
	@PostMapping("/saveUser")
	@Operation(summary = "saveUser", description = "this api will save the user")
	public ResponseEntity<UserDto> saveUser(@RequestBody UserDto dto){
		UserDto saveUser = userService.saveUser(dto);
		return ResponseEntity.status(HttpStatus.CREATED).body(saveUser);
	}
	@GetMapping("/getUserByUsername")
	@CrossOrigin
	@Operation(summary = "getUserByUsername", description = "this api will give the User details")
    public ResponseEntity<UserDto> getUserByUsername(@RequestParam("username")String username) {
        return ResponseEntity.ok(userService.getUserByUsername(username));
    }
	

}
