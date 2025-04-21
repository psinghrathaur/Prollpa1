package com.prollpa.controller;

import java.util.List;


import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import com.prollpa.entity.Role;
import com.prollpa.payload.UserDto;
import com.prollpa.service.RoleService;
import com.prollpa.serviceImpl.RoleServiceImpl;

import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;

@RestController
@RequestMapping("/roles")
@Tag(name = "roles Controller Api", description = "roles Management")
public class RoleController {
	private final RoleService roleService;
	
	public RoleController(RoleService roleService) {
		super();
		this.roleService = roleService;
	}
	
	
	@PostMapping("/saveRole")
	@Operation(summary = "saveRole", description = "this api will save the Role")
	public ResponseEntity<Role> saveRole(@RequestBody Role role){
		 if (role == null) {
		        return ResponseEntity.badRequest().body(null); // Return a 400 response
		    }
		    System.out.println("Received Role: " + role);
		    Role saveRole = roleService.saveRole(role);
		    return ResponseEntity.status(HttpStatus.CREATED).body(saveRole);
}
	@GetMapping("/getRoleList")
	@Operation(summary = "getRoleList", description = "this api will give the Role List")
    public ResponseEntity<List<Role>> getRoleList() {
        return ResponseEntity.ok(roleService.getRoleList());
    }
	@GetMapping("/getRoleByRoleId/{roleId}")
	@Operation(summary = "getRoleByRoleId", description = "this api will give the Role List")
    public ResponseEntity<Role> getRoleByRoleId(@PathVariable("roleId")Long roleId) {
        return ResponseEntity.ok(roleService.getRoleByRoleId(roleId));
    }
	
	
	

}
