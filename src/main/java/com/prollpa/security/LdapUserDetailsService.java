package com.prollpa.security;


import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.ldap.core.AttributesMapper;
import org.springframework.ldap.core.LdapTemplate;
import org.springframework.ldap.filter.EqualsFilter;
import org.springframework.security.core.userdetails.User;
import org.springframework.security.core.userdetails.UserDetails;
import org.springframework.security.core.userdetails.UserDetailsService;
import org.springframework.security.core.userdetails.UsernameNotFoundException;
import org.springframework.stereotype.Service;

import com.prollpa.exception.ResourceNotFoundException;

@Service
public class LdapUserDetailsService implements UserDetailsService {

    @Autowired
    private LdapTemplate ldapTemplate;

//    @Override
//    public UserDetails loadUserByUsernames(String username) throws UsernameNotFoundException {
////    	 if (username == null || !username.equals("validUser")) { // Replace with real LDAP validation logic
////             throw new UsernameNotFoundException("please enter valid username -> you entered wrong " + username);
////         }
//        // Implement LDAP user retrieval logic here
//    	
//
//        return User.builder()
//                .username(username)
//                .password("{noop}password") // No password encoding for now
//                .roles("USER")
//                .build();
//    }
    @Override
    public UserDetails loadUserByUsername(String username) throws UsernameNotFoundException {
    	EqualsFilter filter = new EqualsFilter("uid", username);

        // Query the LDAP directory
        boolean userExists = ldapTemplate.search("", filter.encode(),
                (AttributesMapper<Boolean>) attributes -> true).size() > 0;

        if (!userExists) {
            throw new ResourceNotFoundException("Invalid username: " + username);
        }

        // If the user exists, return user details
        return User.builder()
                .username(username)
                .password("{noop}password") // Replace with real password retrieval
                .roles("USER")
                .build();
    }

    
}

