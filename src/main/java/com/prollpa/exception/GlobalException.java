package com.prollpa.exception;

import java.util.stream.Collectors;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.context.annotation.Configuration;
import org.springframework.dao.InvalidDataAccessApiUsageException;
import org.springframework.http.HttpStatus;


import org.springframework.http.ResponseEntity;
import org.springframework.security.core.AuthenticationException;
import org.springframework.security.core.userdetails.UsernameNotFoundException;
import org.springframework.web.bind.MethodArgumentNotValidException;
import org.springframework.web.bind.annotation.ControllerAdvice;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.ResponseStatus;
import org.springframework.web.bind.annotation.RestControllerAdvice;
import org.springframework.web.context.request.WebRequest;

import com.prollpa.exception.*;


import io.jsonwebtoken.ExpiredJwtException;


@RestControllerAdvice
@Configuration
public class GlobalException {
	private static final Logger logger = LoggerFactory.getLogger(GlobalException.class);
	@ExceptionHandler(UsernameNotFoundException.class)
	public ResponseEntity<ExceptionApiResponse> handleUserNotFoundException(UsernameNotFoundException exception,WebRequest webRequest){
		return ResponseEntity.status(HttpStatus.NOT_FOUND).body(ExceptionApiResponse.builder().status(HttpStatus.NOT_FOUND).message(exception.getMessage()).description(webRequest.getDescription(false)).build());		
	}
	@ExceptionHandler(ExpiredJwtException.class)
    public ResponseEntity<ExceptionApiResponse> handleExpiredJwtException(ExpiredJwtException exception, WebRequest webRequest) {
        ExceptionApiResponse response = ExceptionApiResponse.builder()
                .status(HttpStatus.UNAUTHORIZED)
                .message("Token has expired. Please login again.")
                .description(webRequest.getDescription(false))
                .build();

        return ResponseEntity.status(HttpStatus.UNAUTHORIZED).body(response);
    }

	@ExceptionHandler(Exception.class)
    public ResponseEntity<ExceptionApiResponse> handleException(Exception exception, WebRequest webRequest) {
        ExceptionApiResponse response = ExceptionApiResponse.builder()
                .status(HttpStatus.NOT_FOUND) // Changed to 401 Unauthorized
                .message(exception.getMessage()) // More specific message
                .description(webRequest.getDescription(false))
                .build();

        return ResponseEntity.status(HttpStatus.NOT_FOUND).body(response);
    }
	@ExceptionHandler(ResourceNotFoundException.class)
	public ResponseEntity<ExceptionApiResponse> handleResourceNotFoundException(ResourceNotFoundException exception, WebRequest webRequest) {
	    logger.error("Handling ResourceNotFoundException: {}", exception.getMessage());
	    ExceptionApiResponse response = ExceptionApiResponse.builder()
	            .status(HttpStatus.NOT_FOUND)
	            .message(exception.getMessage())
	            .description(webRequest.getDescription(false))
	            .build();

	    return ResponseEntity.status(HttpStatus.NOT_FOUND).body(response);
	}
	@ExceptionHandler(AuthenticationException.class)
    public ResponseEntity<String> handleAuthenticationException(AuthenticationException e) {
        return ResponseEntity.status(HttpStatus.UNAUTHORIZED).body("Authentication failed: " + e.getMessage());
    }
	
	@ExceptionHandler(MethodArgumentNotValidException.class)
    public ResponseEntity<ExceptionApiResponse> handleValidationExceptions(MethodArgumentNotValidException ex, WebRequest webRequest) {
        // Collect error messages from validation
        String errorMessages = ex.getBindingResult().getAllErrors().stream()
                .map(error -> error.getDefaultMessage())
                .collect(Collectors.joining("; "));

        // Build the custom response
        ExceptionApiResponse response = ExceptionApiResponse.builder()
                .status(HttpStatus.BAD_REQUEST) // 400 Bad Request
                .message(ex.getMessage())
                .description(errorMessages)
                .build();

        // Return the custom response with BAD_REQUEST status
        return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(response);
    }
	@ExceptionHandler(InvalidDataAccessApiUsageException.class)
    public ResponseEntity<ExceptionApiResponse> handleValidationsExceptions(InvalidDataAccessApiUsageException ex, WebRequest webRequest) {
        // Collect error messages from validation

        // Build the custom response
        ExceptionApiResponse response = ExceptionApiResponse.builder()
                .status(HttpStatus.BAD_REQUEST) // 400 Bad Request
                .message(ex.getMessage())
                .description(webRequest.getDescription(false))
                .build();

        // Return the custom response with BAD_REQUEST status
        return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(response);
    }
	
	
	
}
