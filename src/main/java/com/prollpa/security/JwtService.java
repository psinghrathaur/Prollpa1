package com.prollpa.security;
import io.jsonwebtoken.Claims;

import io.jsonwebtoken.Jwts;
import io.jsonwebtoken.SignatureAlgorithm;
import io.jsonwebtoken.io.Decoders;
import io.jsonwebtoken.security.Keys;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.security.core.userdetails.UserDetails;
import org.springframework.stereotype.Component;

import java.security.Key;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.function.Function;



import io.jsonwebtoken.Jwts;
import io.jsonwebtoken.SignatureAlgorithm;
import io.jsonwebtoken.Claims;
import io.jsonwebtoken.security.Keys;
import org.springframework.stereotype.Service;

import javax.crypto.SecretKey;
import java.util.Date;

@Service
public class JwtService {
	@Value("${jwtTokenExpiration}")
	private  int jwtexpirationTime;

    // Secret key for signing the JWT, should be kept secure and not hardcoded
    private static final String SECRET = "5367566B59703373367639792F423F4528482B4D6251655468576D5A71347437"; 
    private static final SecretKey SECRET_KEY = Keys.hmacShaKeyFor(SECRET.getBytes());

    // Token validity time (e.g., 30 minutes)
    private static final long EXPIRATION_TIME = 1000 * 60  ; // 30 minutes

    // Generate JWT token for a given username
    public String generateToken(String username) {
        return Jwts.builder()
                .setSubject(username) // Set the subject (username) as the claim
                .setIssuedAt(new Date()) // Set the issue date
                .setExpiration(new Date(System.currentTimeMillis() + EXPIRATION_TIME*jwtexpirationTime)) // Set the expiration date
                .signWith(SECRET_KEY, SignatureAlgorithm.HS256) // Sign the token with the secret key
                .compact(); // Return the token as a compact string
    }

    // Extract the username from the JWT token
    public String extractUsername(String token) {
        return Jwts.parserBuilder()
                .setSigningKey(SECRET_KEY) // Use the same secret key to validate the token
                .build()
                .parseClaimsJws(token) // Parse the JWT to extract claims
                .getBody()
                .getSubject(); // Extract the subject (username)
    }

    // Validate the JWT token
    public boolean validateToken(String token, String username) {
        final String extractedUsername = extractUsername(token); // Extract username from token
        return (extractedUsername.equals(username)) && !isTokenExpired(token); // Check if the username matches and if the token is not expired
    }

    // Check if the token is expired
    public boolean isTokenExpired(String token) {
        return extractExpiration(token).before(new Date()); // Token is expired if expiration date is before current date
    }

    // Extract the expiration date from the token
    private Date extractExpiration(String token) {
        return Jwts.parserBuilder()
                .setSigningKey(SECRET_KEY)
                .build()
                .parseClaimsJws(token)
                .getBody()
                .getExpiration();
    }
}
