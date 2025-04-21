package com.prollpa.controller;

import com.prollpa.exception.ResourceNotFoundException;
import com.prollpa.payload.LoginDto;
import com.prollpa.payload.LoginResponse;
import com.prollpa.security.JwtService;
import com.prollpa.security.LdapUserDetailsService;

import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;
import jakarta.validation.Valid;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.security.authentication.AuthenticationManager;
import org.springframework.security.authentication.BadCredentialsException;
import org.springframework.security.authentication.UsernamePasswordAuthenticationToken;
import org.springframework.security.core.Authentication;
import org.springframework.security.core.AuthenticationException;
import org.springframework.security.core.userdetails.UserDetails;
import org.springframework.validation.BindingResult;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/auth")
@Tag(name = "Auth Controller API", description = "Authentication Management")
public class AuthController {

    @Autowired
    private AuthenticationManager authenticationManager;

    @Autowired
    private JwtService jwtService;
    @Autowired
    private LdapUserDetailsService ldapService;
    @Value("${jwtTokenExpiration}")
    private int tokenExpiryTime;

    
    @PostMapping("/ldapLogin")
    @Operation(summary = "Authenticate user and generate JWT token")
    public ResponseEntity<?> login(@RequestBody @Valid LoginDto loginDto , BindingResult bindingresult) {
    	
    	UserDetails loadUserByUsername = ldapService.loadUserByUsername(loginDto.getUsername());
        try {
            // Authenticate the user
     
            Authentication authentication = authenticationManager.authenticate(
                new UsernamePasswordAuthenticationToken(loginDto.getUsername(), loginDto.getPassword())
            );

            // Generate JWT token upon successful authentication
            String token = jwtService.generateToken(loginDto.getUsername());
            LoginResponse loginResponse=LoginResponse.builder().jwtToken(token).validateTime(tokenExpiryTime+" Minutes").authenticationStatus("Authentication Successfully!").build();
            
            return ResponseEntity.ok(loginResponse);

        } catch (BadCredentialsException e) {
        	throw new ResourceNotFoundException("Invalid Password: " );
        } catch (AuthenticationException e) {
            return ResponseEntity.status(HttpStatus.UNAUTHORIZED).body("Authentication failed: " + e.getMessage());
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("An error occurred: " + e.getMessage());
        }
        
    }

   
}
