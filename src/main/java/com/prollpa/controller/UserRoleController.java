package com.prollpa.controller;

import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.validation.BindingResult;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.prollpa.entity.Role;
import com.prollpa.exception.ResourceNotFoundException;
import com.prollpa.payload.UserRoleDto;
import com.prollpa.service.UserRoleService;

import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;
import jakarta.validation.Valid;

@RestController
@RequestMapping("/userRoles")
@Tag(name = "User Role Controller API", description = "Roles Management")
public class UserRoleController {

    private static final Logger logger = LoggerFactory.getLogger(UserRoleController.class);

    private final UserRoleService userRoleService;

    public UserRoleController(UserRoleService userRoleService) {
        this.userRoleService = userRoleService;
    }

    @PostMapping("/saveUserRole")
    @Operation(summary = "saveUserRole", description = "This API will save the user role")
    public ResponseEntity<UserRoleDto> saveUserRole(@RequestBody @Valid UserRoleDto userRoleDto, BindingResult bindingResult) {
        logger.info("Entering saveUserRole API with data: {}", userRoleDto);

        if (bindingResult.hasErrors()) {
            logger.error("Validation errors: {}", bindingResult.getAllErrors());
            throw new ResourceNotFoundException( bindingResult.getAllErrors().toString());
        }
            UserRoleDto savedRole = userRoleService.saveUserRole(userRoleDto);
            logger.info("User role saved successfully: {}", savedRole);
            return ResponseEntity.status(HttpStatus.CREATED).body(savedRole);
        
    }

    @GetMapping("/getUserRoleListByUserId/{userId}")
    @Operation(summary = "getUserRoleList", description = "This API will return the role list by user ID")
    public ResponseEntity<List<UserRoleDto>> getUserRoleList(@PathVariable("userId") Long userId) {
        logger.info("Entering getUserRoleList API with userId: {}", userId);
        List<UserRoleDto> userRoles = userRoleService.getUserRoleByUserId(userId);
        logger.info("Roles retrieved successfully for userId {}: {}", userId, userRoles);
        return ResponseEntity.ok(userRoles); 
    }
}
