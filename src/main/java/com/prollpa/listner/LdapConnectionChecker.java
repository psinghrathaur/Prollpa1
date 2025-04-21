package com.prollpa.listner;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.context.event.ApplicationReadyEvent;
import org.springframework.context.event.EventListener;
import org.springframework.ldap.core.LdapTemplate;
import org.springframework.stereotype.Component;

@Component
public class LdapConnectionChecker {

    @Autowired
    private LdapTemplate ldapTemplate;

    @EventListener(ApplicationReadyEvent.class)
    public void checkLdapConnection() {
        try {
            ldapTemplate.list(""); // Simple query to check connection
            System.out.println("LDAP connection successful!");
        } catch (Exception e) {
            System.err.println("Failed to connect to LDAP: " + e.getMessage());
        }
    }
}
