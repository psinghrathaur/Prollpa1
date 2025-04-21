package com.prollpa.security;



import org.springframework.security.core.userdetails.User;
import org.springframework.security.core.userdetails.UserDetails;
import org.springframework.security.core.userdetails.UserDetailsService;
import org.springframework.security.core.userdetails.UsernameNotFoundException;
import org.springframework.stereotype.Service;

import java.util.Collections;

@Service
public class UserInfoService implements UserDetailsService {

    @Override
    public UserDetails loadUserByUsername(String username) throws UsernameNotFoundException {
        // Replace with actual database lookup logic
        if ("testUser".equals(username)) {
            return new User(
                "testUser", 
                "{noop}password", // Use "{noop}" for plain text password in dev only
                Collections.emptyList() // No roles/authorities in this example
            );
        } else {
            throw new UsernameNotFoundException("User not found: " + username);
        }
    }
}

